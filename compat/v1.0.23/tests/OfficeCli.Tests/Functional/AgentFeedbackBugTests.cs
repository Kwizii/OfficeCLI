// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests that prove bugs found during AI Agent usability testing.
/// These tests are expected to FAIL until the bugs are fixed.
///
/// Bug 4 — Adding a "first" type header puts w:titlePg in w:settings instead of w:sectPr,
///          causing OOXML schema validation errors.
/// Bug 6 — Excel formula cells with no cached value report Format["empty"] = true,
///          which is semantically contradictory (the cell has a formula, it's not empty).
/// </summary>
public class AgentFeedbackBugTests_Word : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public AgentFeedbackBugTests_Word()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        BlankDocCreator.Create(_path);
        _handler = new WordHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private WordHandler Reopen()
    {
        _handler.Dispose();
        _handler = new WordHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Bug 4: titlePg schema validation ====================

    /// <summary>
    /// Bug 4: When adding a "first" type header, the code appends TitlePage to Settings
    /// (settings.xml). According to the OOXML spec, w:titlePg is a child of w:sectPr
    /// (SectionProperties), not w:settings. This produces a schema validation error.
    ///
    /// The test creates a blank docx, adds a first-page header, then validates
    /// the document using the OpenXML SDK validator. It should produce zero errors.
    /// </summary>
    [Fact]
    public void AddFirstPageHeader_ShouldNotProduceSchemaValidationErrors()
    {
        // 1. Add a first-page header
        var headerPath = _handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["type"] = "first",
            ["text"] = "First Page Header"
        });
        headerPath.Should().Contain("header");

        // 2. Get the header and verify it was created
        var node = _handler.Get(headerPath);
        node.Should().NotBeNull();

        // 3. Save and validate: dispose handler to flush, then validate the file
        _handler.Dispose();

        using var doc = WordprocessingDocument.Open(_path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();

        // Filter for titlePg-related errors specifically
        var titlePgErrors = errors
            .Where(e => e.Description != null &&
                        (e.Description.Contains("titlePg", StringComparison.OrdinalIgnoreCase) ||
                         e.Description.Contains("TitlePage", StringComparison.OrdinalIgnoreCase) ||
                         (e.Part?.Uri?.ToString().Contains("settings") == true)))
            .ToList();

        // The document should have zero validation errors related to settings.xml
        titlePgErrors.Should().BeEmpty(
            "adding a first-page header should not produce schema validation errors in settings.xml, " +
            "but titlePg was incorrectly placed in w:settings instead of w:sectPr");

        // Re-create handler for Dispose() in teardown
        _handler = new WordHandler(_path, editable: false);
    }

    /// <summary>
    /// Bug 4 (variant): Verify that titlePg is in the correct location (SectionProperties)
    /// after adding a first-page header. The OOXML spec says w:titlePg belongs in w:sectPr.
    /// </summary>
    [Fact]
    public void AddFirstPageHeader_TitlePgShouldBeInSectionProperties()
    {
        // 1. Add a first-page header
        _handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["type"] = "first",
            ["text"] = "First Page Only"
        });

        // 2. Save
        _handler.Dispose();

        // 3. Inspect the raw XML to check where titlePg ended up
        using var doc = WordprocessingDocument.Open(_path, false);
        var body = doc.MainDocumentPart?.Document?.Body;
        body.Should().NotBeNull();

        var sectPr = body!.Elements<SectionProperties>().LastOrDefault();
        sectPr.Should().NotBeNull("document should have a SectionProperties element");

        // titlePg should be a child of SectionProperties
        var titlePgInSectPr = sectPr!.GetFirstChild<TitlePage>();
        titlePgInSectPr.Should().NotBeNull(
            "w:titlePg should be placed in w:sectPr (SectionProperties), not in w:settings, " +
            "because the OOXML spec defines titlePg as a child element of sectPr");

        // Re-create handler for Dispose() in teardown
        _handler = new WordHandler(_path, editable: false);
    }

    // ==================== Bug 3: header indexing starts at wrong number ====================

    /// <summary>
    /// Bug 3: When adding the first header to a blank document, the returned path
    /// should be /header[1], not a higher index like /header[4]. A blank document
    /// that never had headers should start indexing at 1.
    /// </summary>
    [Fact]
    public void AddFirstHeader_ShouldReturnHeaderIndex1()
    {
        // Add the very first header to a blank document
        var headerPath = _handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["text"] = "My Header"
        });

        // The first header added to a blank document should be /header[1]
        headerPath.Should().Be("/header[1]",
            "the first header added to a blank document should be indexed as [1], " +
            "not a higher number that includes internal pre-existing header parts");
    }
}

public class AgentFeedbackBugTests_Excel : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public AgentFeedbackBugTests_Excel()
    {
        _path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        BlankDocCreator.Create(_path);
        _handler = new ExcelHandler(_path, editable: true);
    }

    public void Dispose()
    {
        _handler.Dispose();
        if (File.Exists(_path)) File.Delete(_path);
    }

    private ExcelHandler Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
        return _handler;
    }

    // ==================== Bug 6: Formula cell shows empty: true ====================

    /// <summary>
    /// Bug 6: An Excel cell containing a formula (but no cached value) reports
    /// Format["empty"] = true. This is semantically contradictory — a cell with
    /// a formula is not empty, even if it hasn't been calculated yet.
    ///
    /// The root cause is in ExcelHandler.Helpers.cs CellToNode(): line 285 checks
    /// `string.IsNullOrEmpty(value)` where `value` is the raw cached value, and sets
    /// empty=true regardless of whether the cell has a formula. A formula cell should
    /// never be considered "empty".
    /// </summary>
    [Fact]
    public void FormulaCell_ShouldNotBeMarkedAsEmpty()
    {
        // 1. Add some data cells
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string>
        {
            ["ref"] = "A1",
            ["value"] = "10"
        });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string>
        {
            ["ref"] = "A2",
            ["value"] = "20"
        });

        // 2. Add a formula cell (use "formula" property, not "value")
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string>
        {
            ["ref"] = "A3",
            ["formula"] = "SUM(A1:A2)"
        });

        // 3. Get the formula cell
        var node = _handler.Get("/Sheet1/A3");
        node.Should().NotBeNull();
        node.Type.Should().Be("cell");

        // The cell has a formula — it should report it
        node.Format.Should().ContainKey("formula",
            "cell A3 contains a formula and should report it");

        // THIS IS THE BUG: Format["empty"] should not be true for formula cells
        // A cell with a formula has content — the formula itself. It is not empty.
        node.Format.Should().NotContainKey("empty",
            "a cell with a formula (even uncalculated) is not semantically empty — " +
            "it has content (the formula). Marking it empty:true contradicts having formula content");
    }

    /// <summary>
    /// Bug 6 (variant): Even after reopening the file, a formula cell without a cached
    /// value should not be marked as empty.
    /// </summary>
    [Fact]
    public void FormulaCell_ShouldNotBeMarkedAsEmpty_AfterReopen()
    {
        // 1. Add data + formula
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string>
        {
            ["ref"] = "B1", ["value"] = "100"
        });
        _handler.Add("/Sheet1", "cell", null, new Dictionary<string, string>
        {
            ["ref"] = "B2", ["formula"] = "B1*2"
        });

        // 2. Reopen to test persistence
        Reopen();

        // 3. Get the formula cell
        var node = _handler.Get("/Sheet1/B2");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("formula");

        // Formula cells should not be marked empty
        node.Format.Should().NotContainKey("empty",
            "formula cell B2 should not be marked as empty after reopen — it has formula content");
    }
}
