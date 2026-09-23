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
/// Tests that prove bugs found during Agent A Round 4 testing.
/// These tests are expected to FAIL until the bugs are fixed.
///
/// BUG-1: Add style ignores "basedOn" property (camelCase) because AddStyle uses
///        case-sensitive TryGetValue("basedon") while Set uses key.ToLowerInvariant()
///        which matches any casing. The style is created without inheritance.
///
/// BUG-2: Set style appends StyleParagraphProperties after StyleRunProperties,
///        violating OOXML schema order (pPr must come before rPr in w:style).
///        PowerPoint/Word may silently ignore out-of-order elements.
///
/// BUG-3: Add table ignores "colWidths" (camelCase) because AddTable uses
///        case-sensitive TryGetValue("colwidths"). Only lowercase works.
///        All columns get default equal width instead of specified widths.
///
/// BUG-5: Get with out-of-bounds index on /toc[N], /footnote[N], /endnote[N]
///        returns a DocumentNode with Type="error" instead of throwing.
///        Same class of issue as Round 3 Bug 6 (/section[N] out-of-bounds).
/// </summary>
public class AgentFeedbackBugTests_Round4 : IDisposable
{
    private readonly string _path;
    private WordHandler _handler;

    public AgentFeedbackBugTests_Round4()
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

    // ==================== BUG-1: Add style ignores basedOn (camelCase) ====================

    /// <summary>
    /// BUG-1: When creating a style with basedOn="Normal" using camelCase key,
    /// the basedOn property should be set on the style. Instead, AddStyle only
    /// checks for lowercase "basedon" via case-sensitive TryGetValue, so the
    /// camelCase "basedOn" is silently ignored.
    /// </summary>
    [Fact]
    public void Add_Style_WithBasedOnCamelCase_ShouldSetInheritance()
    {
        // Create a base style first
        _handler.Add("/", "style", null, new()
        {
            ["id"] = "MyBase",
            ["name"] = "My Base Style",
            ["type"] = "paragraph",
            ["bold"] = "true",
            ["size"] = "16"
        });

        // Create a derived style using camelCase "basedOn"
        _handler.Add("/", "style", null, new()
        {
            ["id"] = "MyDerived",
            ["name"] = "My Derived Style",
            ["type"] = "paragraph",
            ["basedOn"] = "MyBase",   // camelCase — this is the canonical form
            ["italic"] = "true"
        });

        // Verify: Get the derived style and check basedOn is set
        var node = _handler.Get("/styles/MyDerived");
        node.Should().NotBeNull();
        node.Type.Should().Be("style");

        // The style should have basedOn = "MyBase"
        // Currently fails: basedOn is silently ignored because AddStyle
        // only checks TryGetValue("basedon") (all lowercase)
        node.Format.Should().ContainKey("basedOn");
        node.Format["basedOn"].ToString().Should().Be("MyBase");
    }

    /// <summary>
    /// BUG-1 (persistence variant): basedOn should survive reopen.
    /// Uses lowercase "basedon" to confirm the feature works at all,
    /// then tests that camelCase produces the same result.
    /// </summary>
    [Fact]
    public void Add_Style_WithBasedOnLowercase_Works_ButCamelCaseDoesNot()
    {
        // Lowercase "basedon" — should work (direct TryGetValue match)
        _handler.Add("/", "style", null, new()
        {
            ["id"] = "LowercaseChild",
            ["name"] = "Lowercase Child",
            ["type"] = "paragraph",
            ["basedon"] = "Normal"
        });

        // CamelCase "basedOn" — should also work but currently doesn't
        _handler.Add("/", "style", null, new()
        {
            ["id"] = "CamelCaseChild",
            ["name"] = "CamelCase Child",
            ["type"] = "paragraph",
            ["basedOn"] = "Normal"
        });

        Reopen();

        // Lowercase version should have basedOn
        var lowNode = _handler.Get("/styles/LowercaseChild");
        lowNode.Format.Should().ContainKey("basedOn");
        lowNode.Format["basedOn"].ToString().Should().Be("Normal");

        // CamelCase version should ALSO have basedOn — currently fails
        var camelNode = _handler.Get("/styles/CamelCaseChild");
        camelNode.Format.Should().ContainKey("basedOn");
        camelNode.Format["basedOn"].ToString().Should().Be("Normal");
    }

    // ==================== BUG-2: Set style pPr inserted after rPr ====================

    /// <summary>
    /// BUG-2: When a style already has StyleRunProperties and then Set adds
    /// StyleParagraphProperties (e.g., alignment), the pPr is appended AFTER rPr.
    /// OOXML schema requires pPr to come before rPr within w:style.
    ///
    /// Schema order: styleName, aliases, basedOn, next, link, autoRedefine, ...,
    ///               pPr, rPr, tblPr, trPr, tcPr, ...
    /// </summary>
    [Fact]
    public void Set_Style_ParagraphProperties_ShouldPrecedeRunProperties()
    {
        // Create style with run properties only (bold, font)
        _handler.Add("/", "style", null, new()
        {
            ["id"] = "TestStyleOrder",
            ["name"] = "Test Style Order",
            ["type"] = "paragraph",
            ["bold"] = "true",
            ["font"] = "Arial",
            ["size"] = "14"
        });

        // Now Set paragraph properties (alignment) — this should insert pPr BEFORE rPr
        _handler.Set("/styles/TestStyleOrder", new()
        {
            ["alignment"] = "center"
        });

        // Access the raw OpenXml style to verify element ordering
        _handler.Dispose();
        var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(_path, false);
        try
        {
            var styles = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            var style = styles?.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == "TestStyleOrder");
            style.Should().NotBeNull("style should exist");

            var pPr = style!.StyleParagraphProperties;
            var rPr = style.StyleRunProperties;
            pPr.Should().NotBeNull("pPr should exist after Set alignment");
            rPr.Should().NotBeNull("rPr should exist from Add");

            // Verify schema order: pPr must come before rPr in the child elements
            var children = style.ChildElements.ToList();
            var pPrIndex = children.IndexOf(pPr!);
            var rPrIndex = children.IndexOf(rPr!);

            pPrIndex.Should().BeLessThan(rPrIndex,
                "OOXML schema requires StyleParagraphProperties (pPr) before StyleRunProperties (rPr). " +
                "Currently Set appends pPr at the end, after rPr.");
        }
        finally
        {
            doc.Dispose();
            _handler = new WordHandler(_path, editable: true);
        }
    }

    /// <summary>
    /// BUG-2 (variant): Adding spacing via Set after style creation with run props
    /// should also maintain correct element order.
    /// </summary>
    [Fact]
    public void Set_Style_SpaceBefore_ShouldPrecedeRunProperties()
    {
        _handler.Add("/", "style", null, new()
        {
            ["id"] = "SpacingOrderTest",
            ["name"] = "Spacing Order Test",
            ["type"] = "paragraph",
            ["color"] = "#FF0000"
        });

        _handler.Set("/styles/SpacingOrderTest", new()
        {
            ["spaceBefore"] = "12pt"
        });

        _handler.Dispose();
        var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(_path, false);
        try
        {
            var style = doc.MainDocumentPart?.StyleDefinitionsPart?.Styles
                ?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == "SpacingOrderTest");
            style.Should().NotBeNull();

            var pPr = style!.StyleParagraphProperties;
            var rPr = style.StyleRunProperties;
            pPr.Should().NotBeNull();
            rPr.Should().NotBeNull();

            var children = style.ChildElements.ToList();
            children.IndexOf(pPr!).Should().BeLessThan(children.IndexOf(rPr!),
                "pPr must precede rPr in schema order");
        }
        finally
        {
            doc.Dispose();
            _handler = new WordHandler(_path, editable: true);
        }
    }

    // ==================== BUG-3: Add table ignores colWidths (camelCase) ====================

    /// <summary>
    /// BUG-3: When creating a table with colWidths="3000,5000,2000" (camelCase),
    /// the column widths should be applied. Instead, AddTable only checks for
    /// lowercase "colwidths" via case-sensitive TryGetValue, so the camelCase
    /// "colWidths" is silently ignored and all columns get default 2400 twips.
    /// </summary>
    [Fact]
    public void Add_Table_WithColWidthsCamelCase_ShouldApplyWidths()
    {
        _handler.Add("/", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3",
            ["colWidths"] = "3000,5000,2000"  // camelCase — should work
        });

        // Get the table and check column widths
        var tblNode = _handler.Get("/body/tbl[1]", depth: 2);
        tblNode.Should().NotBeNull();
        tblNode.Type.Should().Be("table");

        // Reopen to flush, then check via raw XML that GridColumn widths are set correctly
        _handler.Dispose();
        var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(_path, false);
        try
        {
            var body = doc.MainDocumentPart?.Document?.Body;
            var table = body?.Elements<Table>().FirstOrDefault();
            table.Should().NotBeNull();

            var gridCols = table!.Elements<TableGrid>().FirstOrDefault()
                ?.Elements<GridColumn>().ToList();
            gridCols.Should().NotBeNull();
            gridCols.Should().HaveCount(3);

            // Verify that the widths are NOT all default (2400)
            // Currently fails: all three columns are 2400 because "colWidths" doesn't match "colwidths"
            gridCols![0].Width!.Value.Should().Be("3000");
            gridCols[1].Width!.Value.Should().Be("5000");
            gridCols[2].Width!.Value.Should().Be("2000");
        }
        finally
        {
            doc.Dispose();
            _handler = new WordHandler(_path, editable: true); // reopen for Dispose()
        }
    }

    /// <summary>
    /// BUG-3 (cell width variant): When colWidths is applied, each cell's
    /// TableCellWidth should also match the specified column width.
    /// </summary>
    [Fact]
    public void Add_Table_WithColWidthsCamelCase_CellWidthsShouldMatch()
    {
        _handler.Add("/", "table", null, new()
        {
            ["rows"] = "1",
            ["cols"] = "2",
            ["colWidths"] = "4000,6000"  // camelCase
        });

        _handler.Dispose();
        var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(_path, false);
        try
        {
            var table = doc.MainDocumentPart?.Document?.Body?.Elements<Table>().First();
            var firstRow = table!.Elements<TableRow>().First();
            var cells = firstRow.Elements<TableCell>().ToList();
            cells.Should().HaveCount(2);

            // First cell width should be 4000, not default
            var cell1Width = cells[0].TableCellProperties?.TableCellWidth?.Width;
            cell1Width.Should().NotBeNull();
            cell1Width!.Value.Should().Be("4000",
                "cell width should match colWidths specification, not default");

            var cell2Width = cells[1].TableCellProperties?.TableCellWidth?.Width;
            cell2Width.Should().NotBeNull();
            cell2Width!.Value.Should().Be("6000");
        }
        finally
        {
            doc.Dispose();
            _handler = new WordHandler(_path, editable: true); // reopen for Dispose()
        }
    }

    // ==================== BUG-5: Out-of-bounds Get returns error node instead of throwing ====================

    /// <summary>
    /// BUG-5: Get('/toc[99]') when only 1 TOC exists should throw an exception
    /// (like /section[N] does after Round 3 fix), not return a node with Type="error".
    /// Returning Type="error" inside a success envelope is misleading for API consumers.
    /// </summary>
    [Fact]
    public void Get_TocOutOfBounds_ShouldThrow()
    {
        // Add one TOC
        _handler.Add("/", "paragraph", null, new() { ["text"] = "Heading", ["style"] = "Heading1" });
        _handler.Add("/", "toc", null, new() { ["levels"] = "1-3" });

        // Verify TOC 1 exists
        var toc1 = _handler.Get("/toc[1]");
        toc1.Should().NotBeNull();
        toc1.Type.Should().Be("toc");

        // Out-of-bounds should throw, not return Type="error"
        // Currently fails: returns DocumentNode with Type="error" instead of throwing
        var act = () => _handler.Get("/toc[99]");
        act.Should().Throw<ArgumentException>("out-of-bounds TOC index should throw, not return error node");
    }

    /// <summary>
    /// BUG-5 (footnote variant): Get('/footnote[99]') should throw for non-existent footnote.
    /// </summary>
    [Fact]
    public void Get_FootnoteOutOfBounds_ShouldThrow()
    {
        // Add a paragraph with a footnote
        _handler.Add("/", "paragraph", null, new() { ["text"] = "Text" });
        _handler.Add("/body/p[1]", "footnote", null, new() { ["text"] = "A footnote" });

        // Verify footnote 1 exists
        var fn1 = _handler.Get("/footnote[1]");
        fn1.Should().NotBeNull();
        fn1.Type.Should().Be("footnote");

        // Out-of-bounds should throw
        var act = () => _handler.Get("/footnote[99]");
        act.Should().Throw<ArgumentException>("out-of-bounds footnote index should throw, not return error node");
    }

    /// <summary>
    /// BUG-5 (endnote variant): Get('/endnote[99]') should throw for non-existent endnote.
    /// </summary>
    [Fact]
    public void Get_EndnoteOutOfBounds_ShouldThrow()
    {
        // Add a paragraph with an endnote
        _handler.Add("/", "paragraph", null, new() { ["text"] = "Text" });
        _handler.Add("/body/p[1]", "endnote", null, new() { ["text"] = "An endnote" });

        // Verify endnote 1 exists
        var en1 = _handler.Get("/endnote[1]");
        en1.Should().NotBeNull();
        en1.Type.Should().Be("endnote");

        // Out-of-bounds should throw
        var act = () => _handler.Get("/endnote[99]");
        act.Should().Throw<ArgumentException>("out-of-bounds endnote index should throw, not return error node");
    }
}
