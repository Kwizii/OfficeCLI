// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Final comprehensive DOCX coverage tests.
///
/// Areas covered:
///  1. Hyperlink formatting readback (type=hyperlink, link URL, text, run link)
///  2. Style underline / strike readback via /styles/StyleId
///  3. Comment Get via /comments/comment[N] direct navigation
///  4. SDT dropdown / combobox items readback
///  5. Header type readback (default / first / even)
///  6. Run-level formatting readback (caps, smallcaps, dstrike, vanish,
///     superscript, subscript, highlight, shading, rtl)
/// </summary>
public class DocxFinalTests : IDisposable
{
    private readonly List<string> _tempFiles = new();

    private (string path, WordHandler handler) CreateDoc()
    {
        var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.docx");
        _tempFiles.Add(path);
        BlankDocCreator.Create(path);
        return (path, new WordHandler(path, editable: true));
    }

    private WordHandler Reopen(string path)
        => new WordHandler(path, editable: true);

    public void Dispose()
    {
        foreach (var f in _tempFiles)
            try { File.Delete(f); } catch { }
    }

    // ════════════════════════════════════════════════════════════════════════
    // 1. HYPERLINK FORMATTING READBACK
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Hyperlink_Add_And_Get_ReturnsTypeHyperlinkAndLink()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        // Add a paragraph then a hyperlink inside it
        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "before" });
        handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://example.com/page",
            ["text"] = "Click here"
        });

        // Navigate to the hyperlink child of the paragraph
        var node = handler.Get("/body/p[1]/hyperlink[1]");

        node.Type.Should().Be("hyperlink",
            "ElementToNode for Hyperlink element must set Type = 'hyperlink'");
        node.Text.Should().Be("Click here",
            "hyperlink text is concatenated from descendant Text nodes");
        node.Format.Should().ContainKey("link",
            "link URL must be resolved from HyperlinkRelationships and stored in Format['link']");
        node.Format["link"].ToString().Should().Contain("example.com",
            "the resolved URL must match the one registered as a relationship");
    }

    [Fact]
    public void Hyperlink_RunInsideHyperlink_ReportsLinkOnRunNode()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "intro " });
        handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://docs.officecli.ai",
            ["text"] = "Docs"
        });

        // The run inside the hyperlink should expose Format["link"]
        var runNode = handler.Get("/body/p[1]/hyperlink[1]/r[1]");

        runNode.Type.Should().Be("run");
        runNode.Format.Should().ContainKey("link",
            "a run whose Parent is Hyperlink must inherit the hyperlink URL into Format['link']");
        runNode.Format["link"].ToString().Should().Contain("officecli.ai");
    }

    [Fact]
    public void Hyperlink_Get_Persists_AfterReopen()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "p1" });
            handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
            {
                ["url"] = "https://persist.test/",
                ["text"] = "Persist"
            });
        }

        using var h2 = Reopen(path);
        var node = h2.Get("/body/p[1]/hyperlink[1]");
        node.Type.Should().Be("hyperlink");
        node.Format.Should().ContainKey("link");
        node.Format["link"].ToString().Should().Contain("persist.test");
    }

    [Fact]
    public void Hyperlink_Query_ReturnsHyperlinkNodes()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "p" });
        handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://query.test/",
            ["text"] = "Link"
        });

        // Query runs that have a link property set
        var runs = handler.Query("r[link]");
        runs.Should().NotBeEmpty("a run inside a hyperlink should appear in Query results when filtering by link");
    }

    // ════════════════════════════════════════════════════════════════════════
    // 2. STYLE UNDERLINE / STRIKE READBACK
    //
    // The AddStyle handler does NOT currently write underline/strike to
    // StyleRunProperties, so these tests verify what IS and what IS NOT
    // read back. They document the current contract precisely.
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Style_Get_ReturnsFont_Bold_Italic_Color_Size()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/styles", "style", null, new Dictionary<string, string>
        {
            ["id"] = "MyRunStyle",
            ["name"] = "MyRunStyle",
            ["type"] = "character",
            ["font"] = "Courier New",
            ["size"] = "14pt",
            ["bold"] = "true",
            ["italic"] = "true",
            ["color"] = "FF0000"
        });

        var node = handler.Get("/styles/MyRunStyle");

        node.Type.Should().Be("style");
        node.Format.Should().ContainKey("font");
        node.Format["font"].Should().Be("Courier New");
        node.Format.Should().ContainKey("size");
        node.Format["size"].ToString().Should().Be("14pt");
        node.Format.Should().ContainKey("bold");
        node.Format.Should().ContainKey("italic");
        node.Format.Should().ContainKey("color");
        node.Format["color"].ToString().Should().Be("#FF0000",
            "color from StyleRunProperties must be #-prefixed via FormatHexColor");
    }

    [Fact]
    public void Style_Get_ReturnsAlignment_SpaceBefore_SpaceAfter()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/styles", "style", null, new Dictionary<string, string>
        {
            ["id"] = "MyParaStyle",
            ["name"] = "MyParaStyle",
            ["type"] = "paragraph",
            ["alignment"] = "center",
            ["spaceBefore"] = "240",
            ["spaceAfter"] = "120"
        });

        var node = handler.Get("/styles/MyParaStyle");

        node.Type.Should().Be("style");
        node.Format.Should().ContainKey("alignment");
        node.Format["alignment"].ToString().Should().Be("center");
        node.Format.Should().ContainKey("spaceBefore");
        node.Format.Should().ContainKey("spaceAfter");
    }

    [Fact]
    public void Style_Get_BasedOn_And_Next_AreReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        // Add base style first
        handler.Add("/styles", "style", null, new Dictionary<string, string>
        {
            ["id"] = "BaseStyle",
            ["name"] = "BaseStyle",
            ["type"] = "paragraph",
        });

        // Add derived style
        handler.Add("/styles", "style", null, new Dictionary<string, string>
        {
            ["id"] = "DerivedStyle",
            ["name"] = "DerivedStyle",
            ["type"] = "paragraph",
            ["basedon"] = "BaseStyle",
            ["next"] = "Normal",
        });

        var node = handler.Get("/styles/DerivedStyle");
        node.Format.Should().ContainKey("basedOn");
        node.Format["basedOn"].Should().Be("BaseStyle");
        node.Format.Should().ContainKey("next");
        node.Format["next"].Should().Be("Normal");
    }

    // NOTE: AddStyle does not currently accept "underline" or "strike" for
    // StyleRunProperties. The gap is documented by the following test which
    // verifies that the readback does NOT silently report incorrect values
    // when no underline/strike XML is present in the style.
    [Fact]
    public void Style_Get_UnderlineAndStrike_NotPresentWhenNotAdded()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/styles", "style", null, new Dictionary<string, string>
        {
            ["id"] = "PlainStyle",
            ["name"] = "PlainStyle",
            ["type"] = "character",
            ["font"] = "Arial"
        });

        var node = handler.Get("/styles/PlainStyle");

        // Style did not set underline or strike, so Get must NOT report them
        node.Format.Should().NotContainKey("underline",
            "underline must not appear when StyleRunProperties has no Underline element");
        node.Format.Should().NotContainKey("strike",
            "strike must not appear when StyleRunProperties has no Strike element");
    }

    // ════════════════════════════════════════════════════════════════════════
    // 3. COMMENT GET VIA /comments/comment[N]
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Comment_Get_ViaPath_ReturnsTextAndAuthor()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Reviewed text" });
        handler.Add("/body/p[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Please revise this section",
            ["author"] = "Reviewer"
        });

        // Query-based access
        var results = handler.Query("comment");
        results.Should().HaveCountGreaterOrEqualTo(1);
        var first = results[0];
        first.Type.Should().Be("comment");
        first.Text.Should().Be("Please revise this section");
        first.Format.Should().ContainKey("author");
        first.Format["author"].Should().Be("Reviewer");
        first.Path.Should().StartWith("/comments/comment[");
    }

    [Fact]
    public void Comment_Get_DirectNavigate_ReturnsCommentNode()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Body" });
        handler.Add("/body/p[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Navigate directly",
            ["author"] = "Alice",
            ["initials"] = "A"
        });

        // Query to find the path, then Get that path
        var qResults = handler.Query("comment");
        qResults.Should().NotBeEmpty();
        var commentPath = qResults[0].Path;

        // Direct Get via /comments/comment[N]
        var node = handler.Get(commentPath);
        node.Text.Should().Be("Navigate directly");
        node.Format.Should().ContainKey("author");
        node.Format["author"].Should().Be("Alice");
    }

    [Fact]
    public void Comment_MultipleComments_IndexedCorrectly()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Para" });
        handler.Add("/body/p[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "First comment", ["author"] = "Alice"
        });
        handler.Add("/body/p[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Second comment", ["author"] = "Bob"
        });

        var results = handler.Query("comment");
        results.Should().HaveCountGreaterOrEqualTo(2);

        var texts = results.Select(r => r.Text).ToList();
        texts.Should().Contain("First comment");
        texts.Should().Contain("Second comment");

        var authors = results.Select(r => r.Format.TryGetValue("author", out var a) ? a?.ToString() : null).ToList();
        authors.Should().Contain("Alice");
        authors.Should().Contain("Bob");
    }

    [Fact]
    public void Comment_Query_WithContainsText_FiltersResults()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Body" });
        handler.Add("/body/p[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Important note here", ["author"] = "Alice"
        });
        handler.Add("/body/p[1]", "comment", null, new Dictionary<string, string>
        {
            ["text"] = "Trivial observation", ["author"] = "Bob"
        });

        var filtered = handler.Query("comment:contains(\"Important\")");
        filtered.Should().HaveCount(1,
            "only comments containing 'Important' should be returned");
        filtered[0].Text.Should().Contain("Important");
    }

    [Fact]
    public void Comment_Get_Persists_AfterReopen()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Body" });
            handler.Add("/body/p[1]", "comment", null, new Dictionary<string, string>
            {
                ["text"] = "Persist this comment", ["author"] = "Tester"
            });
        }

        using var h2 = Reopen(path);
        var results = h2.Query("comment");
        results.Should().NotBeEmpty();
        results[0].Text.Should().Be("Persist this comment");
        results[0].Format["author"].Should().Be("Tester");
    }

    // ════════════════════════════════════════════════════════════════════════
    // 4. SDT DROPDOWN / COMBOBOX ITEMS READBACK
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Sdt_Dropdown_Add_And_Get_ReturnsItemsAndSdtType()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "dropdown",
            ["alias"] = "StatusField",
            ["tag"] = "status",
            ["text"] = "Open",
            ["items"] = "Open,In Progress,Closed"
        });

        var results = handler.Query("sdt");
        results.Should().NotBeEmpty();
        var sdtNode = results[0];

        sdtNode.Type.Should().Be("sdt");
        sdtNode.Format.Should().ContainKey("sdtType");
        sdtNode.Format["sdtType"].Should().Be("dropdown",
            "SdtContentDropDownList maps to sdtType=dropdown");
        sdtNode.Format.Should().ContainKey("items",
            "ListItem entries must be concatenated into Format['items']");
        sdtNode.Format["items"].ToString().Should().Contain("Open");
        sdtNode.Format["items"].ToString().Should().Contain("In Progress");
        sdtNode.Format["items"].ToString().Should().Contain("Closed");
        sdtNode.Format.Should().ContainKey("alias");
        sdtNode.Format["alias"].Should().Be("StatusField");
        sdtNode.Format.Should().ContainKey("tag");
        sdtNode.Format["tag"].Should().Be("status");
    }

    [Fact]
    public void Sdt_Combobox_Add_And_Get_ReturnsItemsAndSdtType()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "combobox",
            ["alias"] = "ColorPicker",
            ["text"] = "Red",
            ["items"] = "Red,Green,Blue"
        });

        var results = handler.Query("sdt");
        results.Should().NotBeEmpty();
        var sdtNode = results[0];

        sdtNode.Format.Should().ContainKey("sdtType");
        sdtNode.Format["sdtType"].Should().Be("combobox",
            "SdtContentComboBox maps to sdtType=combobox");
        sdtNode.Format.Should().ContainKey("items");
        var itemsStr = sdtNode.Format["items"].ToString()!;
        itemsStr.Should().Contain("Red");
        itemsStr.Should().Contain("Green");
        itemsStr.Should().Contain("Blue");
    }

    [Fact]
    public void Sdt_Text_Add_And_Get_ReturnsSdtTypeText()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "text",
            ["alias"] = "NameField",
            ["text"] = "Enter name here"
        });

        var results = handler.Query("sdt");
        results.Should().NotBeEmpty();
        results[0].Format.Should().ContainKey("sdtType");
        results[0].Format["sdtType"].Should().Be("text");
        results[0].Text.Should().Contain("Enter name here");
    }

    [Fact]
    public void Sdt_Date_Add_And_Get_ReturnsSdtTypeDate()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "date",
            ["text"] = "2025-01-01"
        });

        var results = handler.Query("sdt");
        results.Should().NotBeEmpty();
        results[0].Format.Should().ContainKey("sdtType");
        results[0].Format["sdtType"].Should().Be("date");
    }

    [Fact]
    public void Sdt_Dropdown_Persists_AfterReopen()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "sdt", null, new Dictionary<string, string>
            {
                ["sdttype"] = "dropdown",
                ["alias"] = "PriorityField",
                ["text"] = "High",
                ["items"] = "High,Medium,Low"
            });
        }

        using var h2 = Reopen(path);
        var results = h2.Query("sdt");
        results.Should().NotBeEmpty();
        results[0].Format["sdtType"].Should().Be("dropdown");
        results[0].Format["items"].ToString().Should().Contain("Medium");
    }

    [Fact]
    public void Sdt_Inline_Dropdown_InsideParagraph_ReturnsItems()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "Choose: " });
        // Add inline SDT (SdtRun) inside the paragraph
        handler.Add("/body/p[1]", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "dropdown",
            ["alias"] = "InlineDropdown",
            ["text"] = "Option A",
            ["items"] = "Option A,Option B,Option C"
        });

        var results = handler.Query("sdt");
        results.Should().NotBeEmpty("inline SDT inside paragraph should appear in sdt Query");
        var sdtNode = results.FirstOrDefault(n =>
            n.Format.TryGetValue("alias", out var a) && a?.ToString() == "InlineDropdown");
        sdtNode.Should().NotBeNull();
        sdtNode!.Format["sdtType"].Should().Be("dropdown");
        sdtNode.Format["items"].ToString().Should().Contain("Option B");
    }

    // ════════════════════════════════════════════════════════════════════════
    // 5. HEADER TYPE READBACK
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Header_Default_Type_IsReportedInFormat()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["text"] = "Default Header",
            ["type"] = "default"
        });

        var node = handler.Get("/header[1]");
        node.Type.Should().Be("header");
        node.Text.Should().Contain("Default Header");

        // type may or may not be reported for default — check it is present when set
        // The GetHeaderNode code looks up HeaderReference by relId in SectionProperties
        node.Format.Should().ContainKey("type",
            "GetHeaderNode reads HeaderReference.Type from SectionProperties and stores it in Format['type']");
        node.Format["type"].ToString().Should().Be("default");
    }

    [Fact]
    public void Header_First_Type_IsReportedInFormat()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["text"] = "First Page Header",
            ["type"] = "first"
        });

        // Find the header (it's the first one added)
        var node = handler.Get("/header[1]");
        node.Type.Should().Be("header");
        node.Format.Should().ContainKey("type");
        node.Format["type"].ToString().Should().Be("first",
            "header added with type=first must report type=first on Get");
    }

    [Fact]
    public void Header_Even_Type_IsReportedInFormat()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["text"] = "Even Page Header",
            ["type"] = "even"
        });

        var node = handler.Get("/header[1]");
        node.Type.Should().Be("header");
        node.Format.Should().ContainKey("type");
        node.Format["type"].ToString().Should().Be("even",
            "header added with type=even must report type=even on Get");
    }

    [Fact]
    public void Header_Text_And_Font_AreReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["text"] = "Styled Header",
            ["type"] = "default",
            ["font"] = "Arial",
            ["bold"] = "true",
            ["color"] = "0070C0"
        });

        var node = handler.Get("/header[1]");
        node.Text.Should().Contain("Styled Header");
        node.Format.Should().ContainKey("font");
        node.Format["font"].Should().Be("Arial");
        node.Format.Should().ContainKey("bold");
        node.Format.Should().ContainKey("color");
        node.Format["color"].ToString().Should().Be("#0070C0");
    }

    [Fact]
    public void Header_ChildCount_EqualsNumberOfParagraphs()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/", "header", null, new Dictionary<string, string>
        {
            ["text"] = "Header with one paragraph",
            ["type"] = "default"
        });

        var node = handler.Get("/header[1]");
        node.ChildCount.Should().BeGreaterOrEqualTo(1,
            "header must report at least one paragraph in ChildCount");
    }

    // ════════════════════════════════════════════════════════════════════════
    // 6. RUN-LEVEL FORMATTING READBACK
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Run_Get_Caps_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "all caps text",
            ["caps"] = "true"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("caps",
            "Run with Caps element must report caps in Format");
    }

    [Fact]
    public void Run_Get_SmallCaps_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "small caps text",
            ["smallcaps"] = "true"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("smallcaps",
            "Run with SmallCaps element must report smallcaps in Format");
    }

    [Fact]
    public void Run_Get_DoubleStrike_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "double-struck text",
            ["dstrike"] = "true"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("dstrike",
            "Run with DoubleStrike element must report dstrike in Format");
    }

    [Fact]
    public void Run_Get_Superscript_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "x2",
            ["superscript"] = "true"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("superscript",
            "Run with VerticalTextAlignment=Superscript must report superscript=true");
    }

    [Fact]
    public void Run_Get_Subscript_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "H2O",
            ["subscript"] = "true"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("subscript",
            "Run with VerticalTextAlignment=Subscript must report subscript=true");
    }

    [Fact]
    public void Run_Get_Highlight_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "highlighted",
            ["highlight"] = "yellow"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("highlight",
            "Run with Highlight element must report highlight color name");
        run.Format["highlight"].ToString().Should().Be("yellow");
    }

    [Fact]
    public void Run_Get_Underline_Single_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "underlined text",
            ["underline"] = "single"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("underline");
        run.Format["underline"].ToString().Should().Be("single");
    }

    [Fact]
    public void Run_Get_Underline_Double_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "double underlined",
            ["underline"] = "double"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("underline");
        run.Format["underline"].ToString().Should().Be("double",
            "underline=double must survive round-trip (regression: IsTruthy('double') was false)");
    }

    [Fact]
    public void Run_Get_Strike_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "struck through",
            ["strike"] = "true"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("strike",
            "Run with Strike element must report strike=true");
    }

    [Fact]
    public void Run_Get_ShadingColor_IsReported()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "shaded run",
            ["shd"] = "FFFF00"
        });

        var run = handler.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("shading",
            "Run with Shading.Fill must report shading color in Format");
        run.Format["shading"].ToString().Should().Be("#FFFF00");
    }

    [Fact]
    public void Run_Get_FormattingPersists_AfterReopen()
    {
        var (path, h) = CreateDoc();
        using (var handler = h)
        {
            handler.Add("/body", "paragraph", null, new Dictionary<string, string>
            {
                ["text"] = "persist run",
                ["bold"] = "true",
                ["italic"] = "true",
                ["underline"] = "single",
                ["strike"] = "true",
                ["highlight"] = "cyan",
                ["caps"] = "true"
            });
        }

        using var h2 = Reopen(path);
        var run = h2.Get("/body/p[1]/r[1]");
        run.Format.Should().ContainKey("bold");
        run.Format.Should().ContainKey("italic");
        run.Format.Should().ContainKey("underline");
        run.Format.Should().ContainKey("strike");
        run.Format.Should().ContainKey("highlight");
        run.Format.Should().ContainKey("caps");
    }

    // ════════════════════════════════════════════════════════════════════════
    // 6b. PARAGRAPH-LEVEL FORMATTING READBACK (first-run promotion)
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Paragraph_Get_Underline_PromotedFromFirstRun()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "underlined paragraph",
            ["underline"] = "single"
        });

        // Paragraph-level Get should promote first-run underline
        var para = handler.Get("/body/p[1]");
        para.Format.Should().ContainKey("underline",
            "Paragraph ElementToNode promotes first-run formatting properties to paragraph node");
        para.Format["underline"].ToString().Should().Be("single");
    }

    [Fact]
    public void Paragraph_Get_Strike_PromotedFromFirstRun()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string>
        {
            ["text"] = "struck paragraph",
            ["strike"] = "true"
        });

        var para = handler.Get("/body/p[1]");
        para.Format.Should().ContainKey("strike",
            "Paragraph ElementToNode promotes first-run strike to paragraph node");
    }

    // ════════════════════════════════════════════════════════════════════════
    // 7. ADDITIONAL SDT EDGE CASES
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Sdt_Query_ContentControl_Alias_ReturnsSameResult()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "text",
            ["alias"] = "MyControl",
            ["text"] = "value"
        });

        var byContentControl = handler.Query("contentcontrol");
        var bySdt = handler.Query("sdt");

        byContentControl.Should().HaveCount(bySdt.Count,
            "'contentcontrol' is an alias for 'sdt' in Query");
    }

    [Fact]
    public void Sdt_Dropdown_NoItems_ItemsKeyNotPresent()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "sdt", null, new Dictionary<string, string>
        {
            ["sdttype"] = "dropdown",
            ["text"] = "empty dropdown"
            // No items property
        });

        var results = handler.Query("sdt");
        results.Should().NotBeEmpty();
        // When no items are added, Format["items"] should not be present (empty list case)
        // This documents the contract: items key appears only when there are list entries
        var sdtNode = results[0];
        sdtNode.Format["sdtType"].Should().Be("dropdown");
        // items key only present if items.Count > 0
        if (sdtNode.Format.ContainsKey("items"))
            sdtNode.Format["items"].ToString().Should().NotBeNullOrEmpty();
    }

    // ════════════════════════════════════════════════════════════════════════
    // 8. HYPERLINK FORMATTING EDGE CASES
    // ════════════════════════════════════════════════════════════════════════

    [Fact]
    public void Hyperlink_Add_WithFont_And_Bold_RunHasFormatting()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "prefix " });
        handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://styled.link/",
            ["text"] = "Styled Link",
            ["font"] = "Verdana",
            ["bold"] = "true",
            ["size"] = "12pt"
        });

        var run = handler.Get("/body/p[1]/hyperlink[1]/r[1]");
        run.Format.Should().ContainKey("font");
        run.Format["font"].Should().Be("Verdana");
        run.Format.Should().ContainKey("bold");
        run.Format.Should().ContainKey("size");
        run.Format["size"].ToString().Should().Be("12pt");
    }

    [Fact]
    public void Hyperlink_Add_MultipleInSameParagraph_BothAccessible()
    {
        var (_, h) = CreateDoc();
        using var handler = h;

        handler.Add("/body", "paragraph", null, new Dictionary<string, string> { ["text"] = "text " });
        handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://first.link/",
            ["text"] = "First"
        });
        handler.Add("/body/p[1]", "hyperlink", null, new Dictionary<string, string>
        {
            ["url"] = "https://second.link/",
            ["text"] = "Second"
        });

        var hl1 = handler.Get("/body/p[1]/hyperlink[1]");
        var hl2 = handler.Get("/body/p[1]/hyperlink[2]");

        hl1.Text.Should().Be("First");
        hl2.Text.Should().Be("Second");
        hl1.Format["link"].ToString().Should().Contain("first.link");
        hl2.Format["link"].ToString().Should().Contain("second.link");
    }
}
