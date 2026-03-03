using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text.Json;
using System.Text.RegularExpressions;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage.Pickers;
using BEPDesigner.WinUI.Models;
using BEPDesigner.WinUI.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using WinRT.Interop;
using OV = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;

namespace BEPDesigner.WinUI;

public sealed partial class MainWindow : Window
{
    private readonly ObservableCollection<TopicOption> _topics = [];

    private readonly string _appDataDir;
    private readonly string _persistFilePath;
    private readonly string _generatedDir;
    private string _lastGeneratedPath = string.Empty;
    private string _templateDocxPath = string.Empty;

    private static readonly IReadOnlyList<string> DefaultTopics =
    [
        "Scope of BEP",
        "Overview",
        "Disclaimers and clarifications",
        "Project Information",
        "Project Stakeholders",
        "Roles and Responsibilities",
        "Project Phases/Milestones/Issuances",
        "Meetings, Communication, and Collaboration",
        "Software and IT Resources",
        "Post-Mortem meeting",
        "Reminders of Good Practices",
        "BIM Responsibilities and Scopes by Discipline",
        "Home Page, Modeling History and Autopublishing",
        "Information Exchange & Scheduling",
        "Collaboration Strategy",
        "Procedure for Saving and Archiving Models",
        "Geo-Referencing, Project Origin and North.",
        "Level of Development",
        "Model Nomenclature",
        "Shared Parameters",
        "Linking Revit",
        "Linking DWG",
        "Worksets",
        "Phasing",
        "Levels",
        "Gridlines",
        "Matchlines",
        "Project Units",
        "Dimension Syles and Precision",
        "Sheet Size",
        "Drawing Scale",
        "Font",
        "Documentation for Special Scenarios",
        "Clash Detection Roles and Responsibilities",
        "Clash Detection Phase Schedule and Milestones",
        "Clash Detection Common Data Environment",
        "Clash Detection Lead",
        "ACC Model Coordination Module - Set Up",
        "ACC Model Coordination - Issue Tracking",
        "Appendix A",
        "Appendix B",
        "Appendix C",
        "Appendix D",
        "Appendix E",
        "Syncing, Publishing, Sharing, Consuming in ACC"
    ];

    private static readonly HashSet<string> EssentialTopics = new(StringComparer.OrdinalIgnoreCase)
    {
        "Project Information",
        "Project Stakeholders",
        "Roles and Responsibilities",
        "Information Exchange & Scheduling",
        "Collaboration Strategy",
        "ACC Model Coordination Module - Set Up",
        "ACC Model Coordination - Issue Tracking",
        "Syncing, Publishing, Sharing, Consuming in ACC"
    };

    public MainWindow()
    {
        InitializeComponent();

        _appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BEPDesigner.WinUI");
        _persistFilePath = Path.Combine(_appDataDir, "last_state.json");
        _generatedDir = Path.Combine(AppContext.BaseDirectory, "Generated");

        Directory.CreateDirectory(_appDataDir);
        Directory.CreateDirectory(_generatedDir);

        SeedDefaults();
        LoadPersistedState();
        TrySetDefaultTemplatePath();

        Closed += MainWindow_Closed;
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args) => SavePersistedState();

    private void SeedDefaults()
    {
        ProjectNameBox.Text = "118 Project Avenue";
        AutoPublishBox.Text = "Friday 22:00 EST";
        PackageNamingBox.Text = "<Project>_Shared for <Purpose>";
        CoordinationMeetingBox.Text = "Bi-Weekly Wednesday 4:00pm EDT";
        RevitVersionBox.Text = "2024.2";
        AutoCadVersionBox.Text = "2024";
        Civil3DVersionBox.Text = "2024";
        AutoOpenGeneratedBox.IsOn = true;
        EnableWatermarkBox.IsOn = false;
        WatermarkTextBox.Text = "DRAFT";

        PackageMethodBox.SelectedIndex = 1;
        SharingFrequencyBox.SelectedIndex = 1;

        ResetTopics(DefaultTopics, deselected: null);

        TemplatePathText.Text = "Template: (not selected)";
        GeneratedPathText.Text = "Generated file path will appear here.";
        StatusText.Text = "Ready";
    }

    private static (string GroupTitle, int SubtopicIndent) GetTopicGroup(int indexInDefaultList)
    {
        if (indexInDefaultList <= 10) return ("Section 1 - Project Foundations", indexInDefaultList == 0 ? 0 : 1);
        if (indexInDefaultList <= 32) return ("Section 2 - BIM Modeling Standards", indexInDefaultList == 11 ? 0 : 1);
        if (indexInDefaultList <= 38) return ("Section 3 - Coordination and Clash Detection", indexInDefaultList == 33 ? 0 : 1);
        return ("Section 4 - Appendices and ACC Workflow", indexInDefaultList == 39 ? 0 : 1);
    }

    private void ResetTopics(IEnumerable<string> topicNames, HashSet<string>? deselected)
    {
        _topics.Clear();
        var usedGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var orderedDistinct = topicNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        var defaultOrder = DefaultTopics.ToList();

        foreach (var t in orderedDistinct)
        {
            var defaultIndex = defaultOrder.FindIndex(d => string.Equals(d, t, StringComparison.OrdinalIgnoreCase));
            var (groupTitle, indentLevel) = GetTopicGroup(defaultIndex < 0 ? int.MaxValue : defaultIndex);
            var showGroupHeader = usedGroups.Add(groupTitle);

            _topics.Add(new TopicOption
            {
                Name = t,
                Keep = deselected is null || !deselected.Contains(t),
                GroupTitle = groupTitle,
                ShowGroupHeader = showGroupHeader,
                IndentMargin = new Thickness(indentLevel * 18, 0, 0, 0)
            });
        }

        TopicsListView.ItemsSource = _topics;
        UpdateTopicSelectionText();
    }

    private void UpdateTopicSelectionText()
    {
        var selected = _topics.Count(t => t.Keep);
        TopicSelectionText.Text = $"{selected} selected / {_topics.Count} total";
    }

    private BepPayload BuildPayload() => new()
    {
        ProjectName = ProjectNameBox.Text.Trim(),
        ProjectNumber = ProjectNumberBox.Text.Trim(),
        ProjectAddress = ProjectAddressBox.Text.Trim(),
        Client = ClientBox.Text.Trim(),
        ProjectType = ProjectTypeBox.Text.Trim(),
        ContractType = ContractTypeBox.Text.Trim(),
        ProjectDescription = ProjectDescriptionBox.Text.Trim(),
        BimLead = BimLeadBox.Text.Trim(),
        CoordinationMeetingCadence = CoordinationMeetingBox.Text.Trim(),
        PackageMethod = (PackageMethodBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty,
        AutoPublishCadence = AutoPublishBox.Text.Trim(),
        SharingFrequency = (SharingFrequencyBox.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? string.Empty,
        PackageNamingConvention = PackageNamingBox.Text.Trim(),
        GeoCoordinateSystem = GeoCoordinateSystemBox.Text.Trim(),
        AcquireCoordinatesFromModel = AcquireCoordinatesModelBox.Text.Trim(),
        RevitVersion = RevitVersionBox.Text.Trim(),
        AutoCadVersion = AutoCadVersionBox.Text.Trim(),
        Civil3DVersion = Civil3DVersionBox.Text.Trim(),
        DesktopConnectorVersion = DesktopConnectorVersionBox.Text.Trim(),
        BluebeamVersion = BluebeamVersionBox.Text.Trim(),
        EnableWatermark = EnableWatermarkBox.IsOn,
        WatermarkText = string.IsNullOrWhiteSpace(WatermarkTextBox.Text) ? "DRAFT" : WatermarkTextBox.Text.Trim(),
        StartFresh = false,
        Sessions = BepDefaults.DefaultClashSessions().Select(s => new ClashSessionState
        {
            Name = s.Name,
            DisciplinePair = s.DisciplinePair,
            Keep = true
        }).ToList()
    };

    private void ApplyPayload(BepPayload payload)
    {
        ProjectNameBox.Text = payload.ProjectName;
        ProjectNumberBox.Text = payload.ProjectNumber;
        ProjectAddressBox.Text = payload.ProjectAddress;
        ClientBox.Text = payload.Client;
        ProjectTypeBox.Text = payload.ProjectType;
        ContractTypeBox.Text = payload.ContractType;
        ProjectDescriptionBox.Text = payload.ProjectDescription;
        BimLeadBox.Text = payload.BimLead;
        CoordinationMeetingBox.Text = payload.CoordinationMeetingCadence;

        SelectComboByContent(PackageMethodBox, payload.PackageMethod);
        SelectComboByContent(SharingFrequencyBox, payload.SharingFrequency);

        AutoPublishBox.Text = payload.AutoPublishCadence;
        PackageNamingBox.Text = payload.PackageNamingConvention;
        GeoCoordinateSystemBox.Text = payload.GeoCoordinateSystem;
        AcquireCoordinatesModelBox.Text = payload.AcquireCoordinatesFromModel;
        RevitVersionBox.Text = payload.RevitVersion;
        AutoCadVersionBox.Text = payload.AutoCadVersion;
        Civil3DVersionBox.Text = payload.Civil3DVersion;
        DesktopConnectorVersionBox.Text = payload.DesktopConnectorVersion;
        BluebeamVersionBox.Text = payload.BluebeamVersion;
        EnableWatermarkBox.IsOn = payload.EnableWatermark;
        WatermarkTextBox.Text = string.IsNullOrWhiteSpace(payload.WatermarkText) ? "DRAFT" : payload.WatermarkText;

    }

    private static void SelectComboByContent(ComboBox comboBox, string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return;
        foreach (var item in comboBox.Items.OfType<ComboBoxItem>())
        {
            if (string.Equals(item.Content?.ToString(), value, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = item;
                return;
            }
        }
    }

    private async void GenerateButton_Click(object sender, RoutedEventArgs e)
    {
        var clickedButton = sender as Button;
        if (clickedButton is not null) clickedButton.IsEnabled = false;
        StatusText.Text = "Generating BEP section...";

        try
        {
            var payload = BuildPayload();
            var result = await PythonBridge.GenerateAsync(payload);
            OutputBox.Text = result;

            _lastGeneratedPath = SaveGeneratedMarkdown(payload.ProjectName, result);
            GeneratedPathText.Text = _lastGeneratedPath;

            if (AutoOpenGeneratedBox.IsOn) OpenPath(_lastGeneratedPath);

            SavePersistedState(payload);
            StatusText.Text = $"Generated. Saved to {_lastGeneratedPath}";
        }
        catch (Exception ex)
        {
            OutputBox.Text = ex.ToString();
            StatusText.Text = "Generation failed.";
        }
        finally
        {
            if (clickedButton is not null) clickedButton.IsEnabled = true;
        }
    }

    private async void ChooseTemplateButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".docx");
            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

            var file = await picker.PickSingleFileAsync();
            if (file is null)
            {
                StatusText.Text = "Template selection cancelled.";
                return;
            }

            _templateDocxPath = file.Path;
            TemplatePathText.Text = "Template: " + _templateDocxPath;
            SavePersistedState();
            StatusText.Text = "Template selected.";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to choose template: " + ex.Message;
        }
    }

    private void FillTemplateDocxButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(_templateDocxPath) || !File.Exists(_templateDocxPath))
            {
                StatusText.Text = "Original template DOCX not found. Click Choose Original DOCX first.";
                return;
            }

            var payload = BuildPayload();
            var safeName = string.IsNullOrWhiteSpace(payload.ProjectName) ? "project" : SanitizeFileName(payload.ProjectName);
            var outputPath = Path.Combine(_generatedDir, $"{safeName}_BEP_FILLED_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

            File.Copy(_templateDocxPath, outputPath, true);
            var sectionsToRemove = _topics.Where(t => !t.Keep).Select(t => t.Name).ToList();
            var changed = FillDocxTemplate(outputPath, payload, sectionsToRemove);

            _lastGeneratedPath = outputPath;
            GeneratedPathText.Text = outputPath;

            if (AutoOpenGeneratedBox.IsOn) OpenPath(outputPath);

            SavePersistedState(payload);
            StatusText.Text = $"Filled DOCX created: {outputPath} (fields/sections updated: {changed})";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to fill original DOCX: " + ex.Message;
        }
    }

    private static int FillDocxTemplate(string docxPath, BepPayload payload, IReadOnlyList<string> sectionsToRemove)
    {
        using var doc = WordprocessingDocument.Open(docxPath, true);
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body is null) return 0;

        if (!string.IsNullOrWhiteSpace(payload.ProjectName)) doc.PackageProperties.Title = payload.ProjectName;
        if (!string.IsNullOrWhiteSpace(payload.ProjectDescription)) doc.PackageProperties.Description = payload.ProjectDescription;
        if (!string.IsNullOrWhiteSpace(payload.Client)) doc.PackageProperties.Creator = payload.Client;

        var fields = new List<(string Label, string Value)>
        {
            ("Project Number (SvN Internal):", payload.ProjectNumber),
            ("Project Name:", payload.ProjectName),
            ("Project Address:", payload.ProjectAddress),
            ("Project Owner/Client:", payload.Client),
            ("Project Type:", payload.ProjectType),
            ("Contract Type:", payload.ContractType),
            ("Project Description:", payload.ProjectDescription),
            ("KICK-OFF MEETING DATE:", payload.CoordinationMeetingCadence),
            ("Package Sharing Timeline", payload.SharingFrequency),
            ("ACC models will be set to Auto-publish weekly on:", payload.AutoPublishCadence),
            ("Package Sharing Convention", payload.PackageNamingConvention),
            ("The Geocoordinate system will be:", payload.GeoCoordinateSystem),
            ("Coordinates will be acquired from the following Revit model.", payload.AcquireCoordinatesFromModel),
            ("Autodesk Revit", payload.RevitVersion),
            ("Autodesk AutoCAD", payload.AutoCadVersion),
            ("Autodesk Civil 3D", payload.Civil3DVersion),
            ("Autodesk Desktop Connector", payload.DesktopConnectorVersion),
            ("Bluebeam Revu", payload.BluebeamVersion)
        };

        var changed = 0;
        foreach (var (label, value) in fields)
        {
            if (string.IsNullOrWhiteSpace(value)) continue;
            if (FillByLabel(body, label, value)) changed++;
        }

        changed += ReplaceKnownTemplateDefaults(body, payload);
        changed += ClearSections(body, sectionsToRemove);
        if (payload.EnableWatermark)
        {
            ApplyWatermark(doc, string.IsNullOrWhiteSpace(payload.WatermarkText) ? "DRAFT" : payload.WatermarkText.Trim());
            changed++;
        }

        doc.MainDocumentPart!.Document.Save();
        return changed;
    }

    private static void ApplyWatermark(WordprocessingDocument doc, string watermarkText)
    {
        var mainPart = doc.MainDocumentPart;
        if (mainPart?.Document?.Body is null) return;

        var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings ??= new Settings();
        if (!settingsPart.Settings.Elements<DisplayBackgroundShape>().Any())
        {
            settingsPart.Settings.Append(new DisplayBackgroundShape());
        }
        settingsPart.Settings.Save();

        var sections = mainPart.Document.Body.Elements<SectionProperties>().ToList();
        if (sections.Count == 0)
        {
            var lastSect = mainPart.Document.Body.Descendants<SectionProperties>().LastOrDefault();
            if (lastSect is not null) sections.Add(lastSect);
        }
        if (sections.Count == 0)
        {
            var newSect = new SectionProperties();
            mainPart.Document.Body.Append(newSect);
            sections.Add(newSect);
        }

        for (var i = 0; i < sections.Count; i++)
        {
            var section = sections[i];
            var headerRef = section.Elements<HeaderReference>()
                .FirstOrDefault(h => h.Type?.Value == HeaderFooterValues.Default);

            HeaderPart headerPart;
            if (headerRef is not null && !string.IsNullOrWhiteSpace(headerRef.Id))
            {
                headerPart = (HeaderPart)mainPart.GetPartById(headerRef.Id!);
            }
            else
            {
                headerPart = mainPart.AddNewPart<HeaderPart>();
                var relId = mainPart.GetIdOfPart(headerPart);
                section.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = relId });
            }

            headerPart.Header ??= new Header();
            headerPart.Header.RemoveAllChildren<Paragraph>();
            headerPart.Header.Append(BuildWatermarkParagraph(watermarkText, i + 1));
            headerPart.Header.Save();
        }
    }

    private static Paragraph BuildWatermarkParagraph(string watermarkText, int suffix)
    {
        var shapeType = new V.Shapetype
        {
            Id = "_x0000_t136",
            CoordinateSize = "1600,21600",
            OptionalNumber = 136,
            Adjustment = "10800",
            EdgePath = "m@7,l@8,m@5,21600l@6,21600e"
        };
        shapeType.Append(new V.Formulas(
            new V.Formula { Equation = "sum #0 0 10800" },
            new V.Formula { Equation = "prod #0 2 1" },
            new V.Formula { Equation = "sum 21600 0 @1" },
            new V.Formula { Equation = "sum 0 0 @2" },
            new V.Formula { Equation = "sum 21600 0 @3" },
            new V.Formula { Equation = "if @0 @3 0" },
            new V.Formula { Equation = "if @0 21600 @1" },
            new V.Formula { Equation = "if @0 0 @2" },
            new V.Formula { Equation = "if @0 @4 21600" },
            new V.Formula { Equation = "mid @5 @6" },
            new V.Formula { Equation = "mid @8 @5" },
            new V.Formula { Equation = "mid @7 @8" },
            new V.Formula { Equation = "mid @6 @7" },
            new V.Formula { Equation = "sum @6 0 @5" }
        ));
        shapeType.Append(new V.Path { AllowTextPath = TrueFalseValue.FromBoolean(true), ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" });
        shapeType.Append(new V.TextPath { On = TrueFalseValue.FromBoolean(true), FitShape = TrueFalseValue.FromBoolean(true) });
        shapeType.Append(new OV.Lock { TextLock = TrueFalseValue.FromBoolean(true), ShapeType = TrueFalseValue.FromBoolean(true) });

        var shape = new V.Shape
        {
            Id = $"PowerPlusWaterMarkObject{suffix}",
            Style = "position:absolute;margin-left:0;margin-top:0;width:468pt;height:117pt;z-index:-251654144;mso-wrap-edited:f;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
            Type = "#_x0000_t136",
            FillColor = "#d8d8d8",
            Stroked = TrueFalseValue.FromBoolean(false)
        };
        shape.Append(new V.Fill { Opacity = ".5" });
        shape.Append(new V.TextPath { Style = "font-family:&quot;Calibri&quot;;font-size:1pt", String = watermarkText });

        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Header" }),
            new Run(new Picture(shapeType, shape))
        );
    }

    private static bool FillByLabel(Body body, string label, string value)
    {
        if (FillByTableCell(body, label, value)) return true;

        var paragraphs = body.Descendants<Paragraph>().ToList();
        for (var i = 0; i < paragraphs.Count; i++)
        {
            var p = paragraphs[i];
            var text = (p.InnerText ?? string.Empty).Trim();
            if (text.Length == 0 || !text.Contains(label, StringComparison.OrdinalIgnoreCase)) continue;

            if (text.Equals(label, StringComparison.OrdinalIgnoreCase) || text.StartsWith(label, StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 < paragraphs.Count)
                {
                    var next = paragraphs[i + 1];
                    var nextText = (next.InnerText ?? string.Empty).Trim();
                    if (!string.IsNullOrWhiteSpace(nextText) && nextText.Length < 120)
                    {
                        SetParagraphText(next, value);
                        return true;
                    }
                }

                SetParagraphText(p, label + " " + value);
                return true;
            }

            SetParagraphText(p, label + " " + value);
            return true;
        }

        return false;
    }

    private static bool FillByTableCell(Body body, string label, string value)
    {
        foreach (var row in body.Descendants<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            for (var i = 0; i < cells.Count; i++)
            {
                var cellText = (cells[i].InnerText ?? string.Empty).Trim();
                if (!cellText.Contains(label, StringComparison.OrdinalIgnoreCase)) continue;

                // Some templates embed sample values in the label cell
                // (for example: "Project Name:160 John Street"). Normalize
                // the label cell first to avoid duplicated values.
                if (!cellText.Equals(label, StringComparison.OrdinalIgnoreCase))
                {
                    SetCellText(cells[i], label);
                }

                if (i + 1 < cells.Count)
                {
                    SetCellText(cells[i + 1], value);
                    return true;
                }

                SetCellText(cells[i], label + " " + value);
                return true;
            }
        }

        return false;
    }

    private static void SetParagraphText(Paragraph paragraph, string text)
    {
        var firstText = paragraph.Descendants<Text>().FirstOrDefault();
        if (firstText is not null)
        {
            firstText.Text = text;
            foreach (var extra in paragraph.Descendants<Text>().Skip(1)) extra.Text = string.Empty;
            return;
        }

        paragraph.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static void SetCellText(TableCell cell, string text)
    {
        var firstText = cell.Descendants<Text>().FirstOrDefault();
        if (firstText is not null)
        {
            firstText.Text = text;
            foreach (var extra in cell.Descendants<Text>().Skip(1)) extra.Text = string.Empty;
            return;
        }

        var para = cell.Elements<Paragraph>().FirstOrDefault();
        if (para is null)
        {
            para = new Paragraph();
            cell.Append(para);
        }

        para.Append(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static int ReplaceKnownTemplateDefaults(Body body, BepPayload payload)
    {
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(payload.ProjectAddress)) replacements["4926671149796500Project Address"] = payload.ProjectAddress;

        var changed = 0;
        foreach (var textNode in body.Descendants<Text>())
        {
            var current = textNode.Text;
            var updated = current;
            foreach (var kvp in replacements)
            {
                if (updated.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase))
                {
                    updated = Regex.Replace(updated, Regex.Escape(kvp.Key), kvp.Value, RegexOptions.IgnoreCase);
                }
            }

            if (!string.Equals(current, updated, StringComparison.Ordinal))
            {
                textNode.Text = updated;
                changed++;
            }
        }

        return changed;
    }

    private static int ClearSections(Body body, IReadOnlyList<string> sectionsToRemove)
    {
        if (sectionsToRemove.Count == 0) return 0;

        static string NormalizeHeading(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            var lower = input.ToLowerInvariant();
            lower = Regex.Replace(lower, @"\b\d+(\.\d+)*\b", " ");
            lower = Regex.Replace(lower, @"\b[a-z]\.\d+(\.\d+)*\b", " ");
            lower = Regex.Replace(lower, @"[^a-z0-9]+", " ");
            return Regex.Replace(lower, @"\s+", " ").Trim();
        }

        static bool IsHeadingMatch(string blockText, string heading)
        {
            var normalizedBlock = NormalizeHeading(blockText);
            var normalizedHeading = NormalizeHeading(heading);
            if (string.IsNullOrWhiteSpace(normalizedBlock) || string.IsNullOrWhiteSpace(normalizedHeading)) return false;
            if (normalizedBlock.Equals(normalizedHeading, StringComparison.Ordinal)) return true;
            return normalizedBlock.Contains(normalizedHeading, StringComparison.Ordinal);
        }

        static string? TryMatchHeading(string blockText, IReadOnlyList<string> candidates)
        {
            string? best = null;
            var bestLen = -1;
            foreach (var candidate in candidates)
            {
                if (!IsHeadingMatch(blockText, candidate)) continue;
                if (candidate.Length <= bestLen) continue;
                best = candidate;
                bestLen = candidate.Length;
            }

            return best;
        }

        var removeSet = sectionsToRemove
            .Select(NormalizeHeading)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .ToHashSet(StringComparer.Ordinal);
        var blocks = body.ChildElements.ToList();
        var knownHeadings = DefaultTopics.ToList();

        static string GetBlockText(OpenXmlElement element)
            => string.Join(" ", element.Descendants<Text>().Select(t => t.Text)).Trim();

        var headings = new List<(int Index, string Heading)>();
        for (var i = 0; i < blocks.Count; i++)
        {
            var text = GetBlockText(blocks[i]);
            if (text.Length == 0 || text.Contains("PAGEREF", StringComparison.OrdinalIgnoreCase)) continue;
            if (text.Length > 220) continue;

            var matchedHeading = TryMatchHeading(text, knownHeadings);
            if (matchedHeading is not null) headings.Add((i, matchedHeading));
        }

        if (headings.Count == 0) return 0;

        var indicesToRemove = new HashSet<int>();
        for (var i = 0; i < headings.Count; i++)
        {
            var (start, headingText) = headings[i];
            if (!removeSet.Contains(NormalizeHeading(headingText))) continue;

            var endExclusive = (i + 1 < headings.Count) ? headings[i + 1].Index : blocks.Count;
            for (var idx = start; idx < endExclusive; idx++) indicesToRemove.Add(idx);
        }

        if (indicesToRemove.Count == 0) return 0;
        foreach (var idx in indicesToRemove.OrderByDescending(i => i)) blocks[idx].Remove();
        return indicesToRemove.Count;
    }

    private static void OpenPath(string path)
    {
        if (!File.Exists(path)) return;
        Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
    }

    private string SaveGeneratedMarkdown(string projectName, string text)
    {
        var safeName = string.IsNullOrWhiteSpace(projectName) ? "project" : SanitizeFileName(projectName);
        var path = Path.Combine(_generatedDir, $"{safeName}_BEP_{DateTime.Now:yyyyMMdd_HHmmss}.md");
        File.WriteAllText(path, text);
        return path;
    }

    private static string SanitizeFileName(string input)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var chars = input.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray();
        return new string(chars).Replace(' ', '_');
    }

    private void CopyButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(OutputBox.Text))
        {
            StatusText.Text = "Nothing to copy.";
            return;
        }

        var package = new DataPackage();
        package.SetText(OutputBox.Text);
        Clipboard.SetContent(package);
        StatusText.Text = "Output copied to clipboard.";
    }

    private async void SavePresetButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var picker = new FileSavePicker();
            picker.FileTypeChoices.Add("JSON Preset", [".json"]);
            picker.SuggestedFileName = string.IsNullOrWhiteSpace(ProjectNameBox.Text)
                ? "bep-preset"
                : ProjectNameBox.Text.Replace(' ', '-').ToLowerInvariant() + "-preset";

            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
            var file = await picker.PickSaveFileAsync();
            if (file is null)
            {
                StatusText.Text = "Save preset cancelled.";
                return;
            }

            var json = JsonSerializer.Serialize(BuildPayload(), new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(file.Path, json);
            StatusText.Text = $"Preset saved: {file.Path}";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to save preset: " + ex.Message;
        }
    }

    private async void LoadPresetButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(".json");
            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

            var file = await picker.PickSingleFileAsync();
            if (file is null)
            {
                StatusText.Text = "Load preset cancelled.";
                return;
            }

            var json = await File.ReadAllTextAsync(file.Path);
            var payload = JsonSerializer.Deserialize<BepPayload>(json);
            if (payload is null)
            {
                StatusText.Text = "Preset file is empty or invalid.";
                return;
            }

            ApplyPayload(payload);
            SavePersistedState(payload);
            StatusText.Text = $"Preset loaded: {file.Path}";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to load preset: " + ex.Message;
        }
    }

    private void OpenLastGeneratedButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_lastGeneratedPath) || !File.Exists(_lastGeneratedPath))
        {
            StatusText.Text = "No generated file found yet.";
            return;
        }

        OpenPath(_lastGeneratedPath);
        StatusText.Text = "Opened last generated file.";
    }

    private async void ExportDocxButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var text = OutputBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                StatusText.Text = "No generated output. Click Generate first.";
                return;
            }

            var picker = new FileSavePicker();
            picker.FileTypeChoices.Add("Word Document", [".docx"]);
            picker.SuggestedFileName = string.IsNullOrWhiteSpace(ProjectNameBox.Text)
                ? "BEP-Section"
                : ProjectNameBox.Text.Replace(' ', '-') + "-BEP-Section";

            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
            var file = await picker.PickSaveFileAsync();
            if (file is null)
            {
                StatusText.Text = "Export DOCX cancelled.";
                return;
            }

            ExportOutputToDocx(file.Path, text);
            StatusText.Text = $"DOCX exported: {file.Path}";
        }
        catch (Exception ex)
        {
            StatusText.Text = "Failed to export DOCX: " + ex.Message;
        }
    }

    private static void ExportOutputToDocx(string path, string text)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document();
        var body = new Body();

        foreach (var line in text.Replace("\r", string.Empty).Split('\n'))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                body.Append(new Paragraph(new Run(new Text(string.Empty))));
                continue;
            }

            if (line.StartsWith("## ", StringComparison.Ordinal))
            {
                body.Append(BuildStyledParagraph(line[3..], "Heading1"));
                continue;
            }

            if (line.StartsWith("### ", StringComparison.Ordinal))
            {
                body.Append(BuildStyledParagraph(line[4..], "Heading2"));
                continue;
            }

            body.Append(new Paragraph(new Run(new Text(line))));
        }

        main.Document.Append(body);
        main.Document.Save();
    }

    private static Paragraph BuildStyledParagraph(string text, string styleId)
    {
        var p = new Paragraph();
        p.Append(new ParagraphProperties(new ParagraphStyleId { Val = styleId }));
        p.Append(new Run(new Text(text)));
        return p;
    }

    private void TrySetDefaultTemplatePath()
    {
        if (!string.IsNullOrWhiteSpace(_templateDocxPath) && File.Exists(_templateDocxPath))
        {
            TemplatePathText.Text = "Template: " + _templateDocxPath;
            return;
        }

        var known = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "WWP BIM Execution Plan.docx"));
        if (File.Exists(known))
        {
            _templateDocxPath = known;
            TemplatePathText.Text = "Template: " + _templateDocxPath;
        }
        else
        {
            TemplatePathText.Text = "Template: (not selected)";
        }
    }

    private void LoadPersistedState()
    {
        try
        {
            if (!File.Exists(_persistFilePath)) return;

            var json = File.ReadAllText(_persistFilePath);
            var state = JsonSerializer.Deserialize<PersistedState>(json);
            if (state is null) return;

            ApplyPayload(state.Payload);
            AutoOpenGeneratedBox.IsOn = state.AutoOpenGenerated;
            _lastGeneratedPath = state.LastGeneratedPath ?? string.Empty;
            _templateDocxPath = state.TemplateDocxPath ?? string.Empty;

            ResetTopics(DefaultTopics, state.RemovedTopics.ToHashSet(StringComparer.OrdinalIgnoreCase));

            TemplatePathText.Text = string.IsNullOrWhiteSpace(_templateDocxPath) ? "Template: (not selected)" : "Template: " + _templateDocxPath;
            GeneratedPathText.Text = string.IsNullOrWhiteSpace(_lastGeneratedPath) ? "Generated file path will appear here." : _lastGeneratedPath;
        }
        catch
        {
            // Ignore state load failures.
        }
    }

    private void SavePersistedState(BepPayload? payload = null)
    {
        try
        {
            payload ??= BuildPayload();
            var state = new PersistedState
            {
                Payload = payload,
                AutoOpenGenerated = AutoOpenGeneratedBox.IsOn,
                LastGeneratedPath = _lastGeneratedPath,
                TemplateDocxPath = _templateDocxPath,
                RemovedTopics = _topics.Where(t => !t.Keep).Select(t => t.Name).ToList()
            };

            var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_persistFilePath, json);
        }
        catch
        {
            // Ignore persistence errors.
        }
    }

    private void SelectAllTopicsButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _topics) t.Keep = true;
        TopicsListView.ItemsSource = null;
        TopicsListView.ItemsSource = _topics;
        UpdateTopicSelectionText();
        SavePersistedState();
    }

    private void SelectNoneTopicsButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _topics) t.Keep = false;
        TopicsListView.ItemsSource = null;
        TopicsListView.ItemsSource = _topics;
        UpdateTopicSelectionText();
        SavePersistedState();
    }

    private void KeepEssentialTopicsButton_Click(object sender, RoutedEventArgs e)
    {
        foreach (var t in _topics) t.Keep = EssentialTopics.Contains(t.Name);
        TopicsListView.ItemsSource = null;
        TopicsListView.ItemsSource = _topics;
        UpdateTopicSelectionText();
        SavePersistedState();
    }

    private void ReloadTopicsButton_Click(object sender, RoutedEventArgs e)
    {
        ResetTopics(DefaultTopics, deselected: null);
        SavePersistedState();
    }

    private void TopicCheck_Changed(object sender, RoutedEventArgs e)
    {
        UpdateTopicSelectionText();
        SavePersistedState();
    }

    private sealed class PersistedState
    {
        public BepPayload Payload { get; set; } = new();
        public bool AutoOpenGenerated { get; set; } = true;
        public string LastGeneratedPath { get; set; } = string.Empty;
        public string TemplateDocxPath { get; set; } = string.Empty;
        public List<string> RemovedTopics { get; set; } = [];
    }
}
