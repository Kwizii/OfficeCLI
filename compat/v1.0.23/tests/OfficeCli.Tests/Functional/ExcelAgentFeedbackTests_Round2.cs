// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using FluentAssertions;
using OfficeCli;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests.Functional;

/// <summary>
/// Tests for Round 2 bugs reported by Agent A's Excel exploration.
///
/// CONFIRMED BUG (test expected to FAIL until fixed):
///
///   Bug 4 (HIGH) — Chart title.size rejects pt suffix but Get returns pt suffix
///            Example: Set(path, ["title.size"] = "18pt") throws ArgumentException,
///            but Get returns title.size = "18pt".
///            CLAUDE.md rule: "Font size input is lenient: accepts 14, 14pt, 10.5pt"
///            Root cause: ChartSetter.cs line 52 calls SafeParseDouble(value, "title.size")
///            which calls double.TryParse on the raw value. "18pt" fails TryParse, throwing
///            ArgumentException("Invalid 'title.size' value '18pt'").
///            Meanwhile ChartReader.cs line 37 outputs: $"{titleRp.FontSize.Value / 100.0}pt"
///            This creates an asymmetry: Get output cannot be round-tripped back through Set.
///            Fix: Strip "pt" suffix before parsing in the "size" case of ChartSetter, similar
///            to how other handlers use ParseHelpers for lenient font size input.
/// </summary>
public class ExcelAgentFeedbackTests_Round2 : IDisposable
{
    private readonly string _path;
    private ExcelHandler _handler;

    public ExcelAgentFeedbackTests_Round2()
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

    private void Reopen()
    {
        _handler.Dispose();
        _handler = new ExcelHandler(_path, editable: true);
    }

    // =====================================================================
    // Bug 4: chart title.size rejects pt suffix but Get returns pt suffix
    // =====================================================================

    [Fact]
    public void Bug4_ChartTitleSize_WithPtSuffix_ShouldNotThrow()
    {
        // Create a chart with a title
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30",
            ["title"] = "Test Chart"
        });

        // Set title.size with "pt" suffix — CLAUDE.md says font size input is lenient
        var act = () => _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.size"] = "18pt"
        });

        act.Should().NotThrow(
            "CLAUDE.md mandates lenient font size input: accepts '14', '14pt', '10.5pt'. " +
            "But ChartSetter.SafeParseDouble fails on '18pt' because double.TryParse " +
            "cannot parse strings with unit suffixes.");
    }

    [Fact]
    public void Bug4_ChartTitleSize_WithPtSuffix_SetsCorrectValue()
    {
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30",
            ["title"] = "Test Chart"
        });

        // Set with pt suffix
        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.size"] = "18pt"
        });

        // Get should return "18pt"
        var node = _handler.Get("/Sheet1/chart[1]");
        node.Format.Should().ContainKey("title.size");
        ((string)node.Format["title.size"]).Should().Be("18pt");
    }

    [Fact]
    public void Bug4_ChartTitleSize_WithoutPtSuffix_StillWorks()
    {
        // Bare number input should continue to work (regression guard)
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30",
            ["title"] = "Test Chart"
        });

        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.size"] = "24"
        });

        var node = _handler.Get("/Sheet1/chart[1]");
        node.Format.Should().ContainKey("title.size");
        ((string)node.Format["title.size"]).Should().Be("24pt");
    }

    [Fact]
    public void Bug4_ChartTitleSize_RoundTrip_GetOutputCanBeSetBack()
    {
        // The core of the bug: Get output should be valid Set input
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30",
            ["title"] = "Test Chart"
        });

        // First set with bare number
        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.size"] = "14"
        });

        // Read back — Get returns "14pt"
        var node = _handler.Get("/Sheet1/chart[1]");
        var sizeFromGet = (string)node.Format["title.size"];
        sizeFromGet.Should().Be("14pt");

        // Now feed Get output back into Set — this is the round-trip test
        var act = () => _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.size"] = sizeFromGet  // "14pt" — should not throw
        });

        act.Should().NotThrow(
            "Get returns '14pt' but Set rejects it. Input/output must be symmetrical: " +
            "any value produced by Get should be accepted by Set.");

        // Verify the value is unchanged after round-trip
        node = _handler.Get("/Sheet1/chart[1]");
        ((string)node.Format["title.size"]).Should().Be("14pt");
    }

    [Fact]
    public void Bug4_ChartTitleSize_WithPtSuffix_Persists()
    {
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B,C",
            ["series1"] = "S1:10,20,30",
            ["title"] = "Test Chart"
        });

        _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.size"] = "20pt"
        });

        Reopen();

        var node = _handler.Get("/Sheet1/chart[1]");
        node.Format.Should().ContainKey("title.size");
        ((string)node.Format["title.size"]).Should().Be("20pt",
            "title.size set with pt suffix should persist after reopen");
    }

    [Fact]
    public void Bug4_ChartTitleSize_DecimalWithPtSuffix_Works()
    {
        // CLAUDE.md explicitly lists "10.5pt" as valid input
        _handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
        {
            ["chartType"] = "column",
            ["categories"] = "A,B",
            ["series1"] = "S1:5,10",
            ["title"] = "Decimal Size"
        });

        var act = () => _handler.Set("/Sheet1/chart[1]", new Dictionary<string, string>
        {
            ["title.size"] = "10.5pt"
        });

        act.Should().NotThrow(
            "CLAUDE.md says font size input accepts '10.5pt'");

        var node = _handler.Get("/Sheet1/chart[1]");
        node.Format.Should().ContainKey("title.size");
        ((string)node.Format["title.size"]).Should().Be("10.5pt");
    }
}
