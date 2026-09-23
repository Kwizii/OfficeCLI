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
/// Validate-after-every-operation tests for the Word handler.
/// Each Add or Set is followed by a flush (dispose) + OpenXML validation + reopen,
/// so that schema errors are pinpointed to the exact operation that introduced them.
/// </summary>
public class ValidateAfterEveryOpTests_Word : IDisposable
{
    private readonly List<string> _tempFiles = new();
    private readonly ITestOutputHelper _output;

    public ValidateAfterEveryOpTests_Word(ITestOutputHelper output)
    {
        _output = output;
    }

    private string CreateTemp(string ext = ".docx")
    {
        var path = Path.Combine(Path.GetTempPath(), $"validate_op_{Guid.NewGuid():N}{ext}");
        _tempFiles.Add(path);
        return path;
    }

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    /// <summary>
    /// Validates the DOCX file using OpenXmlValidator (Office2019).
    /// Throws assertion failure if any validation errors are found,
    /// with the operation name in the failure message.
    /// </summary>
    private void AssertValid(string path, string operation)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(doc).ToList();
        foreach (var e in errors)
            _output.WriteLine($"[Step: {operation}] {e.ErrorType}: {e.Description} @ {e.Path?.XPath}");
        errors.Should().BeEmpty($"after step: {operation}");
    }

    /// <summary>
    /// Flush (dispose handler), validate, then reopen and return new handler.
    /// </summary>
    private WordHandler FlushValidateReopen(WordHandler handler, string path, string operation)
    {
        handler.Dispose();
        AssertValid(path, operation);
        return new WordHandler(path, editable: true);
    }

    /// <summary>
    /// Creates a minimal valid PNG image (1x1 pixel, red) as a byte array.
    /// </summary>
    private static byte[] MinimalPng()
    {
        // 1x1 red PNG — valid minimal file
        return new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
            0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, // IDAT chunk
            0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
            0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
            0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, // IEND chunk
            0x44, 0xAE, 0x42, 0x60, 0x82
        };
    }

    // =====================================================================
    // Test: Sequential operations — validate after each step (40+ steps)
    // =====================================================================

    [Fact]
    public void SequentialOperations_ValidateAfterEachStep()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        var h = new WordHandler(path, editable: true);

        // ---- Paragraph and text formatting ----

        // Step 1: Add paragraph with text
        h.Add("/body", "paragraph", null, new() { ["text"] = "Hello World" });
        h = FlushValidateReopen(h, path, "1. Add paragraph with text");

        // Step 2: Add run with bold, italic
        h.Add("/body/p[1]", "run", null, new()
        {
            ["text"] = " Bold Italic",
            ["bold"] = "true",
            ["italic"] = "true"
        });
        h = FlushValidateReopen(h, path, "2. Add run with bold, italic");

        // Step 3: Set run: underline=single, color=#FF0000, size=14pt, font=Arial
        h.Set("/body/p[1]/r[2]", new()
        {
            ["underline"] = "single",
            ["color"] = "#FF0000",
            ["size"] = "14pt",
            ["font"] = "Arial"
        });
        h = FlushValidateReopen(h, path, "3. Set run underline/color/size/font");

        // Step 4: Set paragraph: alignment=center, spaceBefore=12pt, spaceAfter=6pt
        h.Set("/body/p[1]", new()
        {
            ["alignment"] = "center",
            ["spaceBefore"] = "12pt",
            ["spaceAfter"] = "6pt"
        });
        h = FlushValidateReopen(h, path, "4. Set paragraph alignment/spaceBefore/spaceAfter");

        // Step 5: Set paragraph: lineSpacing=1.5x
        h.Set("/body/p[1]", new()
        {
            ["lineSpacing"] = "1.5x"
        });
        h = FlushValidateReopen(h, path, "5. Set paragraph lineSpacing=1.5x");

        // Step 6: Set paragraph: indent=720
        h.Set("/body/p[1]", new()
        {
            ["indent"] = "720"
        });
        h = FlushValidateReopen(h, path, "6. Set paragraph indent=720");

        // Step 7: Add paragraph with strikethrough=true
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Strikethrough text",
            ["strikethrough"] = "true"
        });
        h = FlushValidateReopen(h, path, "7. Add paragraph with strikethrough");

        // ---- Lists ----

        // Step 8: Add paragraph with listStyle=bullet
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Bullet item",
            ["listStyle"] = "bullet"
        });
        h = FlushValidateReopen(h, path, "8. Add paragraph listStyle=bullet");

        // Step 9: Add paragraph with listStyle=number, listLevel=1
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Numbered sub-item",
            ["listStyle"] = "number",
            ["listLevel"] = "1"
        });
        h = FlushValidateReopen(h, path, "9. Add paragraph listStyle=number, listLevel=1");

        // ---- Table ----

        // Step 10: Add table (3 rows, 4 cols, colWidths=1000,2000,2000,2200)
        h.Add("/body", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "4",
            ["colWidths"] = "1000,2000,2000,2200"
        });
        h = FlushValidateReopen(h, path, "10. Add table 3x4 with colWidths");

        // Step 11: Set table cell text
        h.Set("/body/tbl[1]/tr[1]/tc[1]", new() { ["text"] = "Header 1" });
        h = FlushValidateReopen(h, path, "11. Set table cell text");

        // Step 12: Set table cell: bold, fill, alignment
        h.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["bold"] = "true",
            ["fill"] = "4472C4",
            ["alignment"] = "center"
        });
        h = FlushValidateReopen(h, path, "12. Set table cell bold/fill/alignment");

        // Step 13: Set table cell: colspan=2
        h.Set("/body/tbl[1]/tr[2]/tc[1]", new()
        {
            ["colspan"] = "2"
        });
        h = FlushValidateReopen(h, path, "13. Set table cell colspan=2");

        // Step 14: Set table cell: border
        h.Set("/body/tbl[1]/tr[1]/tc[2]", new()
        {
            ["border.top"] = "single;8;FF0000"
        });
        h = FlushValidateReopen(h, path, "14. Set table cell border.top");

        // ---- Page setup ----

        // Step 15: Set /: pageWidth=21cm, pageHeight=29.7cm
        h.Set("/", new()
        {
            ["pageWidth"] = "21cm",
            ["pageHeight"] = "29.7cm"
        });
        h = FlushValidateReopen(h, path, "15. Set page width=21cm, height=29.7cm");

        // Step 16: Set /: marginTop=720, marginBottom=720
        h.Set("/", new()
        {
            ["marginTop"] = "720",
            ["marginBottom"] = "720"
        });
        h = FlushValidateReopen(h, path, "16. Set marginTop/marginBottom=720");

        // ---- Header/Footer ----

        // Step 17: Add header (type=default)
        h.Add("/", "header", null, new()
        {
            ["type"] = "default",
            ["text"] = "Default Header"
        });
        h = FlushValidateReopen(h, path, "17. Add header type=default");

        // Step 18: Add footer (type=default)
        h.Add("/", "footer", null, new()
        {
            ["type"] = "default",
            ["text"] = "Default Footer"
        });
        h = FlushValidateReopen(h, path, "18. Add footer type=default");

        // Step 19: Add header (type=first) — titlePg bug regression
        h.Add("/", "header", null, new()
        {
            ["type"] = "first",
            ["text"] = "First Page Header"
        });
        h = FlushValidateReopen(h, path, "19. Add header type=first (titlePg)");

        // ---- Styles ----

        // Step 20: Add style (id=Custom1, name=Custom, basedOn=Normal)
        h.Add("/styles", "style", null, new()
        {
            ["id"] = "Custom1",
            ["name"] = "Custom",
            ["basedOn"] = "Normal"
        });
        h = FlushValidateReopen(h, path, "20. Add style Custom1 basedOn=Normal");

        // Step 21: Set style: font=Arial, size=12pt, bold=true
        h.Set("/styles/Custom1", new()
        {
            ["font"] = "Arial",
            ["size"] = "12pt",
            ["bold"] = "true"
        });
        h = FlushValidateReopen(h, path, "21. Set style font/size/bold");

        // Step 22: Set style: alignment=center, spaceBefore=6pt
        h.Set("/styles/Custom1", new()
        {
            ["alignment"] = "center",
            ["spaceBefore"] = "6pt"
        });
        h = FlushValidateReopen(h, path, "22. Set style alignment/spaceBefore");

        // ---- Hyperlinks and bookmarks ----

        // Step 23: Add paragraph with text (for hyperlink/bookmark targets)
        h.Add("/body", "paragraph", null, new() { ["text"] = "Link target paragraph" });
        h = FlushValidateReopen(h, path, "23. Add paragraph for hyperlink/bookmark");

        // Step 24: Add hyperlink to paragraph
        h.Add("/body/p[5]", "hyperlink", null, new()
        {
            ["url"] = "https://example.com",
            ["text"] = "Click here"
        });
        h = FlushValidateReopen(h, path, "24. Add hyperlink to paragraph");

        // Step 25: Add bookmark
        h.Add("/body/p[5]", "bookmark", null, new()
        {
            ["name"] = "myBookmark",
            ["text"] = "bookmarked text"
        });
        h = FlushValidateReopen(h, path, "25. Add bookmark");

        // ---- Footnotes and endnotes ----

        // Step 26: Add footnote
        h.Add("/body/p[1]", "footnote", null, new()
        {
            ["text"] = "This is a footnote"
        });
        h = FlushValidateReopen(h, path, "26. Add footnote");

        // Step 27: Add endnote
        h.Add("/body/p[1]", "endnote", null, new()
        {
            ["text"] = "This is an endnote"
        });
        h = FlushValidateReopen(h, path, "27. Add endnote");

        // ---- TOC and fields ----

        // Step 28: Add TOC
        h.Add("/body", "toc", null, new()
        {
            ["levels"] = "1-3",
            ["title"] = "Table of Contents"
        });
        h = FlushValidateReopen(h, path, "28. Add TOC");

        // Step 29: Add field (pagenum) to paragraph
        h.Add("/body/p[1]", "field", null, new()
        {
            ["fieldType"] = "pagenum"
        });
        h = FlushValidateReopen(h, path, "29. Add field pagenum");

        // ---- Picture ----

        // Step 30: Create temporary PNG image
        var pngPath = CreateTemp(".png");
        File.WriteAllBytes(pngPath, MinimalPng());

        // Step 31: Add picture
        h.Add("/body", "picture", null, new()
        {
            ["path"] = pngPath,
            ["width"] = "5cm",
            ["height"] = "5cm"
        });
        h = FlushValidateReopen(h, path, "31. Add picture");

        // ---- Section breaks ----

        // Step 32: Add section (type=nextPage)
        h.Add("/body", "section", null, new()
        {
            ["type"] = "nextPage"
        });
        h = FlushValidateReopen(h, path, "32. Add section type=nextPage");

        // Step 33: Set section[2]: orientation=landscape
        h.Set("/section[2]", new()
        {
            ["orientation"] = "landscape"
        });
        h = FlushValidateReopen(h, path, "33. Set section[2] orientation=landscape");

        // Step 34: Set section[2]: pageWidth=29.7cm, pageHeight=21cm
        h.Set("/section[2]", new()
        {
            ["pageWidth"] = "29.7cm",
            ["pageHeight"] = "21cm"
        });
        h = FlushValidateReopen(h, path, "34. Set section[2] pageWidth/pageHeight");

        // ---- Comments and watermark ----

        // Step 35: Add comment to paragraph
        h.Add("/body/p[1]", "comment", null, new()
        {
            ["text"] = "Review this paragraph",
            ["author"] = "TestUser"
        });
        h = FlushValidateReopen(h, path, "35. Add comment to paragraph");

        // Step 36: Add watermark
        h.Add("/body", "watermark", null, new()
        {
            ["text"] = "DRAFT",
            ["color"] = "silver",
            ["font"] = "Calibri"
        });
        h = FlushValidateReopen(h, path, "36. Add watermark");

        // ---- Complex combinations ----

        // Step 37: Add paragraph, set font+color+bold+italic+underline simultaneously
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Rich formatted text",
            ["font"] = "Times New Roman",
            ["color"] = "#0000FF",
            ["bold"] = "true",
            ["italic"] = "true",
            ["underline"] = "single",
            ["size"] = "16pt"
        });
        h = FlushValidateReopen(h, path, "37. Add paragraph with font+color+bold+italic+underline+size");

        // Step 38: Add table, set multiple cell formats
        h.Add("/body", "table", null, new()
        {
            ["rows"] = "2",
            ["cols"] = "3",
            ["colWidths"] = "2000,3000,2000"
        });
        h = FlushValidateReopen(h, path, "38a. Add second table 2x3");

        h.Set("/body/tbl[2]/tr[1]/tc[1]", new()
        {
            ["text"] = "Cell A1",
            ["bold"] = "true",
            ["fill"] = "FF6347"
        });
        h = FlushValidateReopen(h, path, "38b. Set tbl[2] cell A1 text/bold/fill");

        h.Set("/body/tbl[2]/tr[1]/tc[2]", new()
        {
            ["text"] = "Cell B1",
            ["italic"] = "true",
            ["alignment"] = "right"
        });
        h = FlushValidateReopen(h, path, "38c. Set tbl[2] cell B1 text/italic/alignment");

        h.Set("/body/tbl[2]/tr[2]/tc[1]", new()
        {
            ["text"] = "Cell A2",
            ["fill"] = "4682B4",
            ["color"] = "FFFFFF"
        });
        h = FlushValidateReopen(h, path, "38d. Set tbl[2] cell A2 text/fill/color");

        // Final verification: document has content
        var root = h.Get("/", depth: 1);
        root.Should().NotBeNull();
        root!.Children.Should().NotBeEmpty("document should have children after all operations");

        h.Dispose();
    }

    // =====================================================================
    // Test 2: Combined property Set — multiple properties in one call
    // =====================================================================

    [Fact]
    public void CombinedPropertySet_ValidateAfterEachCombo()
    {
        var path = CreateTemp();
        BlankDocCreator.Create(path);
        var h = new WordHandler(path, editable: true);

        // Setup: add paragraph with text
        h.Add("/body", "paragraph", null, new() { ["text"] = "Multi-Set Test" });
        h = FlushValidateReopen(h, path, "Setup: add paragraph");

        // Combo 1: text formatting in one call
        h.Set("/body/p[1]/r[1]", new()
        {
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "FF0000",
            ["size"] = "18pt",
            ["font"] = "Georgia",
            ["underline"] = "single"
        });
        h = FlushValidateReopen(h, path, "Combo 1: bold+italic+color+size+font+underline");

        // Combo 2: paragraph formatting in one call
        h.Set("/body/p[1]", new()
        {
            ["alignment"] = "center",
            ["spaceBefore"] = "12pt",
            ["spaceAfter"] = "6pt",
            ["lineSpacing"] = "2x"
        });
        h = FlushValidateReopen(h, path, "Combo 2: alignment+spaceBefore+spaceAfter+lineSpacing");

        // Combo 3: Add table and set multiple properties at once
        h.Add("/body", "table", null, new()
        {
            ["rows"] = "3",
            ["cols"] = "3",
            ["colWidths"] = "2000,2000,2000"
        });
        h = FlushValidateReopen(h, path, "Combo 3a: add table");

        h.Set("/body/tbl[1]/tr[1]/tc[1]", new()
        {
            ["text"] = "Bold Header",
            ["bold"] = "true",
            ["fill"] = "4472C4",
            ["alignment"] = "center"
        });
        h = FlushValidateReopen(h, path, "Combo 3b: table cell text+bold+fill+alignment");

        // Combo 4: Add style with all properties at once
        h.Add("/styles", "style", null, new()
        {
            ["id"] = "ComboStyle",
            ["name"] = "ComboStyle",
            ["basedOn"] = "Normal",
            ["font"] = "Verdana",
            ["size"] = "11pt",
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "333333"
        });
        h = FlushValidateReopen(h, path, "Combo 4: add style with all run properties");

        // Combo 5: page setup — all margins in one call
        h.Set("/", new()
        {
            ["pageWidth"] = "21cm",
            ["pageHeight"] = "29.7cm",
            ["marginTop"] = "1440",
            ["marginBottom"] = "1440"
        });
        h = FlushValidateReopen(h, path, "Combo 5: pageWidth+pageHeight+marginTop+marginBottom");

        // Combo 6: Add paragraph with list + formatting
        h.Add("/body", "paragraph", null, new()
        {
            ["text"] = "Formatted list item",
            ["listStyle"] = "bullet",
            ["bold"] = "true",
            ["color"] = "008000"
        });
        h = FlushValidateReopen(h, path, "Combo 6: paragraph with listStyle+bold+color");

        // Combo 7: Header with formatting
        h.Add("/", "header", null, new()
        {
            ["type"] = "default",
            ["text"] = "Styled Header",
            ["font"] = "Arial",
            ["size"] = "10pt",
            ["bold"] = "true",
            ["color"] = "666666"
        });
        h = FlushValidateReopen(h, path, "Combo 7: header with text+font+size+bold+color");

        // Combo 8: Footer with formatting
        h.Add("/", "footer", null, new()
        {
            ["type"] = "default",
            ["text"] = "Styled Footer",
            ["font"] = "Arial",
            ["size"] = "9pt",
            ["italic"] = "true"
        });
        h = FlushValidateReopen(h, path, "Combo 8: footer with text+font+size+italic");

        h.Dispose();
    }
}
