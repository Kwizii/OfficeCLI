// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;
using Xunit.Abstractions;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Schema order audit tests: validate that child elements are inserted in correct
/// OOXML schema order. Uses OpenXmlValidator to detect ordering violations.
///
/// High-risk areas identified:
/// 1. Word SectionProperties (pgSz/pgMar/cols ordering with headerReference)
/// 2. Word RunProperties in field creation (rFonts/b/color/sz ordering)
/// 3. PPT Connector Outline children (fill/prstDash/headEnd/tailEnd ordering)
/// 4. Chart Axis elements (title/numFmt/gridlines/majorUnit ordering)
/// 5. Word Settings children (displayBackgroundShape ordering)
/// </summary>
public class SchemaOrderAuditTests : IDisposable
{
    private readonly string _docxPath = Path.Combine(Path.GetTempPath(), $"schema_audit_{Guid.NewGuid():N}.docx");
    private readonly string _pptxPath = Path.Combine(Path.GetTempPath(), $"schema_audit_{Guid.NewGuid():N}.pptx");
    private readonly string _xlsxPath = Path.Combine(Path.GetTempPath(), $"schema_audit_{Guid.NewGuid():N}.xlsx");
    private readonly ITestOutputHelper _output;

    public SchemaOrderAuditTests(ITestOutputHelper output)
    {
        _output = output;
        BlankDocCreator.Create(_docxPath);
        BlankDocCreator.Create(_pptxPath);
        BlankDocCreator.Create(_xlsxPath);
    }

    public void Dispose()
    {
        foreach (var p in new[] { _docxPath, _pptxPath, _xlsxPath })
            try { File.Delete(p); } catch { }
    }

    private List<ValidationErrorInfo> ValidateDocx(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[DOCX] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        return errors;
    }

    private List<ValidationErrorInfo> ValidatePptx(string path)
    {
        using var doc = PresentationDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[PPTX] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        return errors;
    }

    private List<ValidationErrorInfo> ValidateXlsx(string path)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[XLSX] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        return errors;
    }

    // ==================== 1. Word SectionProperties ordering ====================

    /// <summary>
    /// Add header first, then set columns on section.
    /// CT_SectPr order: headerReference* -> ... -> pgSz -> pgMar -> ... -> cols -> ...
    /// Risk: EnsureColumns may place cols before headerReference via PrependChild fallback.
    /// </summary>
    [Fact]
    public void Word_SectPr_AddHeader_ThenSetColumns_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            // Add a paragraph so section has content
            handler.Add("/", "paragraph", null, new() { ["text"] = "Hello" });
            // Add a default header — this inserts HeaderReference into SectionProperties
            handler.Add("/", "header", null, new() { ["text"] = "My Header" });
            // Now set columns on the section — EnsureColumns must place cols AFTER pgSz/pgMar
            handler.Set("/section[1]", new() { ["columns"] = "2" });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("SectionProperties child elements must follow OOXML schema order");
    }

    /// <summary>
    /// Add header + footer, then set margins on section.
    /// Risk: EnsureSectPrPageMargin may use PrependChild when no pgSz exists,
    /// placing PageMargin before HeaderReference/FooterReference.
    /// </summary>
    [Fact]
    public void Word_SectPr_AddHeaderFooter_ThenSetMargins_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            handler.Add("/", "paragraph", null, new() { ["text"] = "Content" });
            handler.Add("/", "header", null, new() { ["text"] = "Header Text" });
            handler.Add("/", "footer", null, new() { ["text"] = "Footer Text" });
            // Set margins — should insert PageMargin after header/footer references
            handler.Set("/section[1]", new() { ["marginTop"] = "1440", ["marginBottom"] = "1440" });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("PageMargin must appear after headerReference/footerReference in SectionProperties");
    }

    /// <summary>
    /// Set pageSize + margins + columns together after adding header.
    /// Comprehensive test: all three Ensure* methods invoked with existing HeaderReference.
    /// </summary>
    [Fact]
    public void Word_SectPr_AllProperties_AfterHeader_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            handler.Add("/", "paragraph", null, new() { ["text"] = "Body" });
            handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "First Page" });
            handler.Add("/", "header", null, new() { ["text"] = "Default Header" });
            handler.Add("/", "footer", null, new() { ["text"] = "Page Footer" });
            // Set all section properties
            handler.Set("/section[1]", new()
            {
                ["pageWidth"] = "12240",
                ["pageHeight"] = "15840",
                ["marginTop"] = "1440",
                ["marginBottom"] = "1440",
                ["marginLeft"] = "1440",
                ["marginRight"] = "1440",
                ["columns"] = "3"
            });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("Combined pageSize + pageMargin + columns must follow schema order after headerReference/footerReference");
    }

    /// <summary>
    /// Add first-page header which appends TitlePage to SectionProperties.
    /// CT_SectPr order places titlePg near the end. Verify it doesn't break schema.
    /// </summary>
    [Fact]
    public void Word_SectPr_FirstPageHeader_TitlePage_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            handler.Add("/", "paragraph", null, new() { ["text"] = "Content" });
            // Add first-page header (triggers TitlePage insertion)
            handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "First Header" });
            // Then set page size and columns to stress-test ordering
            handler.Set("/section[1]", new()
            {
                ["pageWidth"] = "12240",
                ["columns"] = "2"
            });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("TitlePage element must be in correct position within SectionProperties");
    }

    // ==================== 2. Word RunProperties in field creation ====================

    /// <summary>
    /// Create a field (e.g. page number) with font, size, bold, and color.
    /// CT_RPr order: rFonts -> b -> ... -> color -> sz -> ...
    /// Current code: AppendChild(RunFonts), AppendChild(FontSize), AppendChild(Bold), AppendChild(Color)
    /// This puts FontSize before Bold and Color, violating schema order (sz should come after color).
    /// </summary>
    [Fact]
    public void Word_Field_RunProperties_WithAllFormatting_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            handler.Add("/", "field", null, new()
            {
                ["type"] = "PAGE",
                ["font"] = "Arial",
                ["size"] = "12",
                ["bold"] = "true",
                ["color"] = "FF0000"
            });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("RunProperties children in field runs must follow CT_RPr schema order: rFonts, b, color, sz");
    }

    /// <summary>
    /// Create a date field with font + size (no bold/color) to test the simpler case.
    /// CT_RPr: rFonts must come before sz.
    /// </summary>
    [Fact]
    public void Word_DateField_RunProperties_FontAndSize_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            handler.Add("/", "field", null, new()
            {
                ["type"] = "DATE",
                ["font"] = "Times New Roman",
                ["size"] = "10"
            });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("RunProperties with rFonts + sz must follow schema order");
    }

    // ==================== 3. PPT Connector Outline child ordering ====================

    /// <summary>
    /// Create connector, set lineColor + lineDash + tailEnd.
    /// CT_LineProperties order: fill (solidFill) -> prstDash -> ... -> headEnd -> tailEnd
    /// Risk: PrependChild(SolidFill) after prstDash already exists could misorder;
    /// AppendChild(HeadEnd/TailEnd) after prstDash could also misorder.
    /// </summary>
    [Fact]
    public void Pptx_Connector_LineColor_LineDash_TailEnd_SchemaOrderValid()
    {
        using (var handler = new PowerPointHandler(_pptxPath, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "connector", null, new()
            {
                ["x"] = "100000",
                ["y"] = "100000",
                ["width"] = "3000000",
                ["height"] = "0"
            });
            // Set multiple outline properties that need correct ordering
            handler.Set("/slide[1]/connector[1]", new() { ["lineColor"] = "FF0000" });
            handler.Set("/slide[1]/connector[1]", new() { ["lineDash"] = "dash" });
            handler.Set("/slide[1]/connector[1]", new() { ["tailEnd"] = "arrow" });
        }

        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("Connector outline children must follow CT_LineProperties order: solidFill, prstDash, headEnd, tailEnd");
    }

    /// <summary>
    /// Set lineDash first, then lineColor (which uses PrependChild for SolidFill).
    /// The PrependChild should place fill before prstDash, which is correct —
    /// but if headEnd/tailEnd already exist, the ordering may still break.
    /// </summary>
    [Fact]
    public void Pptx_Connector_AllOutlineProps_SchemaOrderValid()
    {
        using (var handler = new PowerPointHandler(_pptxPath, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "connector", null, new()
            {
                ["x"] = "200000",
                ["y"] = "200000",
                ["width"] = "4000000",
                ["height"] = "1000000"
            });
            // Set tailEnd first, then dash, then color — stress tests ordering
            handler.Set("/slide[1]/connector[1]", new() { ["tailEnd"] = "arrow" });
            handler.Set("/slide[1]/connector[1]", new() { ["headEnd"] = "diamond" });
            handler.Set("/slide[1]/connector[1]", new() { ["lineDash"] = "dot" });
            handler.Set("/slide[1]/connector[1]", new() { ["lineColor"] = "0000FF" });
        }

        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("Connector outline must maintain fill -> prstDash -> headEnd -> tailEnd order regardless of Set call order");
    }

    /// <summary>
    /// Set lineColor + lineOpacity on connector.
    /// lineOpacity auto-creates a SolidFill via PrependChild if none exists.
    /// Then set lineDash — prstDash should come after solidFill.
    /// </summary>
    [Fact]
    public void Pptx_Connector_LineOpacity_ThenDash_SchemaOrderValid()
    {
        using (var handler = new PowerPointHandler(_pptxPath, editable: true))
        {
            handler.Add("/", "slide", null, new());
            handler.Add("/slide[1]", "connector", null, new()
            {
                ["x"] = "0",
                ["y"] = "0",
                ["width"] = "5000000",
                ["height"] = "0"
            });
            handler.Set("/slide[1]/connector[1]", new() { ["lineOpacity"] = "0.5" });
            handler.Set("/slide[1]/connector[1]", new() { ["lineDash"] = "longDash" });
            handler.Set("/slide[1]/connector[1]", new() { ["tailEnd"] = "triangle" });
        }

        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("Outline children must maintain schema order after lineOpacity auto-creates SolidFill");
    }

    // ==================== 4. Chart Axis element ordering ====================

    /// <summary>
    /// Set axis title + min + max + gridlines + numberformat on a chart.
    /// CT_ValAx order: axId, scaling, delete, axPos, majorGridlines, minorGridlines,
    ///   title, numFmt, majorTickMark, ...
    /// Risk: multiple InsertAfter/AppendChild calls may produce out-of-order elements.
    /// </summary>
    [Fact]
    public void Excel_Chart_AxisProperties_SchemaOrderValid()
    {
        using (var handler = new ExcelHandler(_xlsxPath, editable: true))
        {
            handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Cat" });
            handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "10" });
            handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B2", ["value"] = "20" });
            handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B3", ["value"] = "30" });
            var chartPath = handler.Add("/Sheet1", "chart", null, new()
            {
                ["chartType"] = "column",
                ["title"] = "Test Chart",
                ["data"] = "S1:10,20,30",
                ["categories"] = "A,B,C"
            });
            // Set multiple axis properties
            handler.Set(chartPath, new() { ["axisTitle"] = "Values" });
            handler.Set(chartPath, new() { ["axisMin"] = "0" });
            handler.Set(chartPath, new() { ["axisMax"] = "50" });
            handler.Set(chartPath, new() { ["gridlines"] = "CCCCCC:0.5" });
            handler.Set(chartPath, new() { ["axisNumFmt"] = "#,##0" });
        }

        var errors = ValidateXlsx(_xlsxPath);
        errors.Should().BeEmpty("ValueAxis children must follow CT_ValAx schema order");
    }

    /// <summary>
    /// Set axis title + gridlines + minor gridlines + major unit on a chart.
    /// Tests that title is placed after gridlines (not before) per schema.
    /// </summary>
    [Fact]
    public void Pptx_Chart_AxisTitle_Gridlines_MajorUnit_SchemaOrderValid()
    {
        using (var handler = new PowerPointHandler(_pptxPath, editable: true))
        {
            handler.Add("/", "slide", null, new());
            var chartPath = handler.Add("/slide[1]", "chart", null, new()
            {
                ["chartType"] = "line",
                ["title"] = "Line Chart",
                ["data"] = "S1:10,20,30;S2:15,25,35",
                ["categories"] = "X,Y,Z"
            });
            handler.Set(chartPath, new() { ["gridlines"] = "true" });
            handler.Set(chartPath, new() { ["minorGridlines"] = "true" });
            handler.Set(chartPath, new() { ["axisTitle"] = "Y-Axis" });
            handler.Set(chartPath, new() { ["majorUnit"] = "10" });
        }

        var errors = ValidatePptx(_pptxPath);
        errors.Should().BeEmpty("Chart axis elements must maintain schema order across multiple Set calls");
    }

    /// <summary>
    /// Set category axis title + value axis numberformat together.
    /// Tests that both axis types maintain their internal ordering.
    /// </summary>
    [Fact]
    public void Excel_Chart_CatTitle_AxisNumFmt_SchemaOrderValid()
    {
        using (var handler = new ExcelHandler(_xlsxPath, editable: true))
        {
            handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "A1", ["value"] = "Cat" });
            handler.Add("/Sheet1", "cell", null, new() { ["ref"] = "B1", ["value"] = "10" });
            var chartPath = handler.Add("/Sheet1", "chart", null, new()
            {
                ["chartType"] = "column",
                ["title"] = "Test",
                ["data"] = "S1:10,20,30",
                ["categories"] = "A,B,C"
            });
            handler.Set(chartPath, new() { ["catTitle"] = "Categories" });
            handler.Set(chartPath, new() { ["axisNumFmt"] = "0.00" });
            handler.Set(chartPath, new() { ["axisMin"] = "0" });
            handler.Set(chartPath, new() { ["axisMax"] = "100" });
        }

        var errors = ValidateXlsx(_xlsxPath);
        errors.Should().BeEmpty("Both CategoryAxis and ValueAxis must maintain schema-ordered children");
    }

    // ==================== 5. Word Settings child ordering ====================

    /// <summary>
    /// Set page background, which appends DisplayBackgroundShape to Settings.
    /// CT_Settings has a specific child order; AppendChild may place it incorrectly
    /// if other settings children already exist.
    /// </summary>
    [Fact]
    public void Word_Settings_PageBackground_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            handler.Add("/", "paragraph", null, new() { ["text"] = "Content" });
            handler.Set("/", new() { ["pageBackground"] = "FFFF00" });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("DisplayBackgroundShape must be in correct position within Settings");
    }

    /// <summary>
    /// Set page background after adding a first-page header (which may also touch Settings).
    /// Tests interaction between header type=first (TitlePage in sectPr) and background (Settings).
    /// </summary>
    [Fact]
    public void Word_Settings_BackgroundAfterFirstHeader_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            handler.Add("/", "paragraph", null, new() { ["text"] = "Body text" });
            handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "First Page" });
            handler.Set("/", new() { ["pageBackground"] = "E0E0E0" });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("Settings must maintain schema order after both first-page header and background operations");
    }

    // ==================== Combined stress tests ====================

    /// <summary>
    /// Full Word document stress test: header + footer + all section properties + field with formatting.
    /// Validates the entire document for any schema ordering violations.
    /// </summary>
    [Fact]
    public void Word_FullDocument_AllFeatures_SchemaOrderValid()
    {
        using (var handler = new WordHandler(_docxPath, editable: true))
        {
            // Add content
            handler.Add("/", "paragraph", null, new() { ["text"] = "Chapter 1" });

            // Add header and footer
            handler.Add("/", "header", null, new() { ["type"] = "first", ["text"] = "First Header" });
            handler.Add("/", "header", null, new() { ["text"] = "Default Header" });
            handler.Add("/", "footer", null, new() { ["text"] = "Page " });

            // Add formatted field
            handler.Add("/", "field", null, new()
            {
                ["type"] = "PAGE",
                ["font"] = "Calibri",
                ["size"] = "11",
                ["bold"] = "true",
                ["color"] = "333333"
            });

            // Set section properties
            handler.Set("/section[1]", new()
            {
                ["pageWidth"] = "12240",
                ["pageHeight"] = "15840",
                ["marginTop"] = "1440",
                ["marginBottom"] = "1440",
                ["marginLeft"] = "1800",
                ["marginRight"] = "1800",
                ["columns"] = "2"
            });

            // Set page background
            handler.Set("/", new() { ["pageBackground"] = "FAFAFA" });
        }

        var errors = ValidateDocx(_docxPath);
        errors.Should().BeEmpty("Full document with headers, footers, fields, sections, and background must pass schema validation");
    }
}
