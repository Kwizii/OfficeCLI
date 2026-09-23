// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests that prove bugs found during Agent A Round 5 testing.
/// These tests are expected to FAIL until the bugs are fixed.
///
/// BUG-A: `colspan` is listed as a valid cell property in the error message
///        (WordHandler.Set.cs line ~1292), but the cell Set switch only handles
///        `case "gridspan":` — there is no `case "colspan":` alias. So
///        `set /body/tbl[1]/tr[1]/tc[1] --prop colspan=2` falls through to
///        the default case and reports UNSUPPORTED.
///        Root cause: Missing "colspan" alias in the cell Set switch statement
///        (WordHandler.Set.cs ~line 1243). The PPT handler does handle "colspan"
///        (PowerPointHandler.ShapeProperties.cs ~line 935) but Word does not.
///
/// BUG-B: `indent` is listed as a valid paragraph property in the error message
///        (WordHandler.Set.cs line ~1002), but neither Set nor Add handles plain
///        "indent" for paragraphs. Set's ApplyParagraphLevelProperty only handles
///        "leftindent"/"indentleft", "rightindent"/"indentright", "firstlineindent",
///        "hangingindent". Add (WordHandler.Add.Text.cs) similarly only handles those
///        specific variants. Plain "indent" should be an alias for "leftindent".
///        Root cause: No `case "indent":` in ApplyParagraphLevelProperty() and no
///        "indent" key check in the Add paragraph properties section.
///        Note: `indent` IS handled for tables (table indentation), but not paragraphs.
/// </summary>
public class AgentFeedbackBugTests_Round5 : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public AgentFeedbackBugTests_Round5()
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

    // ==================== BUG-A: colspan alias missing in Word cell Set ====================

    /// <summary>
    /// BUG-A: Set colspan on a table cell should work since it is listed as a valid
    /// cell property. Currently throws/reports UNSUPPORTED because only "gridspan"
    /// is handled, not "colspan".
    /// </summary>
    [Fact]
    public void Set_TableCell_Colspan_ShouldMergeCells()
    {
        // Create a 3-column table
        _handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3"
        });

        // Set colspan=2 on first cell — should merge first two columns
        // Currently fails: "colspan" is not recognized, falls through to default → UNSUPPORTED
        var act = () => _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["colspan"] = "2"
        });

        // Should not throw or report unsupported
        act.Should().NotThrow();

        // Verify the cell now spans 2 columns
        var cell = _handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Should().NotBeNull();
        cell.Format.Should().ContainKey("gridSpan");
        cell.Format["gridSpan"].ToString().Should().Be("2");
    }

    /// <summary>
    /// BUG-A (control test): "gridspan" works correctly. This test should PASS,
    /// proving the feature exists but the "colspan" alias is missing.
    /// </summary>
    [Fact]
    public void Set_TableCell_GridSpan_Works_AsControl()
    {
        _handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3"
        });

        // "gridspan" is the handled key — this should work
        var act = () => _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["gridspan"] = "2"
        });

        act.Should().NotThrow();

        var cell = _handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Should().NotBeNull();
        cell.Format.Should().ContainKey("gridSpan");
        cell.Format["gridSpan"].ToString().Should().Be("2");
    }

    /// <summary>
    /// BUG-A (persistence): colspan should survive reopen.
    /// </summary>
    [Fact]
    public void Set_TableCell_Colspan_PersistsAfterReopen()
    {
        _handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3"
        });

        _handler.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["colspan"] = "2"
        });

        Reopen();

        var cell = _handler.Get("/body/tbl[1]/tr[1]/tc[1]");
        cell.Should().NotBeNull();
        cell.Format.Should().ContainKey("gridSpan");
        cell.Format["gridSpan"].ToString().Should().Be("2");
    }

    // ==================== BUG-B: paragraph indent not recognized ====================

    /// <summary>
    /// BUG-B (Add): Adding a paragraph with indent=720 should set the left indent.
    /// Currently, "indent" is silently ignored during Add because only "leftindent"
    /// and "indentleft" are checked.
    /// </summary>
    [Fact]
    public void Add_Paragraph_WithIndent_ShouldSetLeftIndent()
    {
        // Add paragraph with plain "indent" property
        _handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Indented text",
            ["indent"] = "720"
        });

        // Verify: Get should return leftIndent = 720
        var node = _handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("leftIndent");
        node.Format["leftIndent"].ToString().Should().Be("720");
    }

    /// <summary>
    /// BUG-B (Set): Setting indent on an existing paragraph should update leftIndent.
    /// Currently reports UNSUPPORTED because ApplyParagraphLevelProperty has no
    /// `case "indent":` — only "leftindent"/"indentleft".
    /// </summary>
    [Fact]
    public void Set_Paragraph_Indent_ShouldUpdateLeftIndent()
    {
        _handler.Add("/", "paragraph", null, new() { ["text"] = "Some text" });

        // Set indent on the paragraph — should set left indent
        var act = () => _handler.Set("/body/p[1]", new()
        {
            ["indent"] = "1440"
        });

        act.Should().NotThrow();

        var node = _handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("leftIndent");
        node.Format["leftIndent"].ToString().Should().Be("1440");
    }

    /// <summary>
    /// BUG-B (control test): "leftindent" works correctly. This test should PASS,
    /// proving the feature exists but the "indent" alias is missing.
    /// </summary>
    [Fact]
    public void Set_Paragraph_LeftIndent_Works_AsControl()
    {
        _handler.Add("/", "paragraph", null, new() { ["text"] = "Some text" });

        var act = () => _handler.Set("/body/p[1]", new()
        {
            ["leftindent"] = "1440"
        });

        act.Should().NotThrow();

        var node = _handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("leftIndent");
        node.Format["leftIndent"].ToString().Should().Be("1440");
    }

    /// <summary>
    /// BUG-B (Add control): "leftindent" works in Add. This should PASS.
    /// </summary>
    [Fact]
    public void Add_Paragraph_WithLeftIndent_Works_AsControl()
    {
        _handler.Add("/", "paragraph", null, new()
        {
            ["text"] = "Indented text",
            ["leftindent"] = "720"
        });

        var node = _handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("leftIndent");
        node.Format["leftIndent"].ToString().Should().Be("720");
    }

    /// <summary>
    /// BUG-B (persistence): indent set via plain "indent" should survive reopen.
    /// </summary>
    [Fact]
    public void Set_Paragraph_Indent_PersistsAfterReopen()
    {
        _handler.Add("/", "paragraph", null, new() { ["text"] = "Some text" });

        _handler.Set("/body/p[1]", new()
        {
            ["indent"] = "1440"
        });

        Reopen();

        var node = _handler.Get("/body/p[1]");
        node.Should().NotBeNull();
        node.Format.Should().ContainKey("leftIndent");
        node.Format["leftIndent"].ToString().Should().Be("1440");
    }
}
