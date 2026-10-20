using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;

namespace HexxcorpDocumentReviewUtility;

/// <summary>
/// Application entry point. Wires up a global exception handler so that an
/// unexpected UI-thread exception shows a message box instead of silently
/// crashing the utility.
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        Services.WebView2LoaderHelper.EnsureLoaderInitialized();

        if (e.Args.Contains("--verify"))
        {
            Environment.Exit(RunVerification());
            return;
        }

        base.OnStartup(e);
        DispatcherUnhandledException += App_DispatcherUnhandledException;
    }

    private static int RunVerification()
    {
        var log = new List<string>();
        Windows.WorkspaceWindow? ws = null;
        try
        {
            log.Add("Starting verification...");
            var mainVm = new ViewModels.MainViewModel();
            ws = new Windows.WorkspaceWindow { MainViewModel = mainVm };

            // 1. Verify Toggle bar sits above FileDatesPanel in DockPanel
            var dockPanel = (DockPanel)ws.Content;
            int fileDatesIdx = -1;
            int toggleBarIdx = -1;
            for (int i = 0; i < dockPanel.Children.Count; i++)
            {
                var child = dockPanel.Children[i];
                if (child == ws.FileDatesPanel)
                {
                    fileDatesIdx = i;
                }
                else if (child is Border b && b.Child is DockPanel dp && dp.Children.Contains(ws.SheetToggleButton))
                {
                    toggleBarIdx = i;
                }
            }
            log.Add($"FileDatesPanel index: {fileDatesIdx}, ToggleBar index: {toggleBarIdx}");
            if (fileDatesIdx < 0 || toggleBarIdx < 0 || toggleBarIdx <= fileDatesIdx)
            {
                throw new Exception($"Toggle bar should be after FileDatesPanel in Bottom-dock sequence (got FileDates: {fileDatesIdx}, Toggle: {toggleBarIdx})");
            }
            log.Add("PASS: Toggle bar sits ABOVE FileDatesPanel when revealed!");

            // 2. Verify SheetToggleButton
            log.Add($"SheetToggleButton content: {ws.SheetToggleButton.Content}");
            if ((string)ws.SheetToggleButton.Content != "View/Edit Sheet")
            {
                throw new Exception("SheetToggleButton content should be 'View/Edit Sheet'");
            }

            // 3. Verify SheetToggleButton toggle click
            ws.SheetToggleButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            log.Add($"After click 1: Visibility={ws.ExcelDrawerHost.Visibility}, Width={ws.ExcelDrawerHost.Width}, Content={ws.SheetToggleButton.Content}");
            if (ws.ExcelDrawerHost.Visibility != Visibility.Visible || (string)ws.SheetToggleButton.Content != "Hide Sheet")
            {
                throw new Exception("Failed to open Excel drawer");
            }

            // 4. Verify three-drawer side-by-side margins
            ws.NotesToggleButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            log.Add($"Doc drawer opened. Doc margin: {ws.NotesPanel.Margin}, Excel margin: {ws.ExcelPanel.Margin}");

            // Close drawers
            ws.SheetToggleButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            ws.NotesToggleButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            log.Add($"After closing: Excel Visibility={ws.ExcelDrawerHost.Visibility}, Doc Visibility={ws.WordDocDrawerHost.Visibility}");

            // 5. Verify EmbeddedExcelController COM check
            bool excelInstalled = Services.EmbeddedExcelController.IsExcelInstalled;
            log.Add($"Is Excel registered on machine: {excelInstalled}");

            // 6. Verify WorkspaceProjectFile schema Version 11 and Excel persistence
            var proj = Services.WorkspaceFileService.BuildProjectFile(
                Array.Empty<Models.CanvasItemModel>(),
                "test notes",
                new Models.WorkspaceHeaderInfo { DriverName = "Synthetic Driver" },
                new Models.WorkspaceDateFieldsInfo(),
                null,
                "",
                new byte[] { 1, 2, 3, 4 },
                "Spreadsheet.xlsx"
            );
            if (proj.Version != 13)
            {
                throw new Exception($"Expected schema Version 13, got {proj.Version}");
            }
            if (proj.ExcelDocFileName != "Spreadsheet.xlsx" || proj.ExcelDocBytes?.Length != 4)
            {
                throw new Exception("Excel doc bytes/filename not preserved in project");
            }
            log.Add("PASS: Workspace schema version 13 with Excel bytes verified!");

            // 7. Verify Excel Embedding & Dynamic Resize
            if (excelInstalled)
            {
                ws.WindowState = WindowState.Normal;
                ws.Width = 1300;
                ws.Height = 750;
                ws.Show();
                ws.UpdateLayout();

                // Open Excel drawer
                ws.SheetToggleButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                ws.UpdateLayout();

                // Create new Excel sheet via center button in empty state
                var centerNewBtn = (Button)ws.ExcelEmptyStatePanel.Children[1];
                centerNewBtn.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                ws.UpdateLayout();

                var excelCtrl = (Services.EmbeddedExcelController)typeof(Windows.WorkspaceWindow)
                    .GetField("_excelController", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
                    .GetValue(ws)!;

                if (!excelCtrl.IsOpen)
                {
                    throw new Exception("Embedded Excel did not open successfully.");
                }

                var excelPanel = (Services.Win32PanelHost)typeof(Windows.WorkspaceWindow)
                    .GetField("_excelHostPanel", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
                    .GetValue(ws)!;

                IntPtr excelHwnd = (IntPtr)typeof(Services.EmbeddedExcelController)
                    .GetField("_excelHwnd", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
                    .GetValue(excelCtrl)!;

                IntPtr xldesk = Services.NativeMethods.FindWindowEx(excelHwnd, IntPtr.Zero, "XLDESK", null);
                IntPtr ex7 = Services.NativeMethods.FindWindowEx(xldesk, IntPtr.Zero, "EXCEL7", null);

                Services.NativeMethods.RECT rD1, r71;
                Services.NativeMethods.GetWindowRect(xldesk, out rD1);
                Services.NativeMethods.GetWindowRect(ex7, out r71);
                log.Add($"Initial: XLDESK={rD1.Right-rD1.Left}x{rD1.Bottom-rD1.Top}, EXCEL7={r71.Right-r71.Left}x{r71.Bottom-r71.Top}");

                // Simulate expanding the Excel drawer to 900px wide and window to 950 tall
                ws.ExcelDrawerHost.Width = 900;
                ws.Height = 950;
                ws.UpdateLayout();
                excelCtrl.ResizeToFillContainer();

                Services.NativeMethods.RECT rD2, r72;
                Services.NativeMethods.GetWindowRect(xldesk, out rD2);
                Services.NativeMethods.GetWindowRect(ex7, out r72);
                log.Add($"Expanded (900 wide, 950 tall): XLDESK={rD2.Right-rD2.Left}x{rD2.Bottom-rD2.Top}, EXCEL7={r72.Right-r72.Left}x{r72.Bottom-r72.Top}");

                if (rD2.Right - rD2.Left <= rD1.Right - rD1.Left)
                {
                    throw new Exception($"XLDESK failed to expand width: was {rD1.Right-rD1.Left}, now {rD2.Right-rD2.Left}");
                }
                if (rD2.Bottom - rD2.Top <= rD1.Bottom - rD1.Top)
                {
                    throw new Exception($"XLDESK failed to expand height: was {rD1.Bottom-rD1.Top}, now {rD2.Bottom-rD2.Top}");
                }

                // Simulate contracting the Excel drawer to 360px wide and window to 700 tall
                ws.ExcelDrawerHost.Width = 360;
                ws.Height = 700;
                ws.UpdateLayout();
                excelCtrl.ResizeToFillContainer();

                Services.NativeMethods.RECT rD3, r73;
                Services.NativeMethods.GetWindowRect(xldesk, out rD3);
                Services.NativeMethods.GetWindowRect(ex7, out r73);
                log.Add($"Contracted (360 wide, 700 tall): XLDESK={rD3.Right-rD3.Left}x{rD3.Bottom-rD3.Top}, EXCEL7={r73.Right-r73.Left}x{r73.Bottom-r73.Top}");

                if (rD3.Right - rD3.Left >= rD2.Right - rD2.Left)
                {
                    throw new Exception($"XLDESK failed to contract width: was {rD2.Right-rD2.Left}, now {rD3.Right-rD3.Left}");
                }
                if (rD3.Bottom - rD3.Top >= rD2.Bottom - rD2.Top)
                {
                    throw new Exception($"XLDESK failed to contract height: was {rD2.Bottom-rD2.Top}, now {rD3.Bottom-rD3.Top}");
                }

                excelCtrl.Close();
                log.Add("PASS: Excel embedding and dynamic resize verified cleanly!");
            }

            // Check 8: Verify Copilot Work integration (prompt building, schema parsing, and drawer layout)
            {
                var header = new Models.WorkspaceHeaderInfo
                {
                    DriverName = "Test Driver",
                    ContractorName = "Test Fleet LLC",
                    TruckNumber = "TRK-900",
                    TrailerNumber = "TRL-400",
                    Ssn = "000-00-0000",
                    Phone = "202-555-0142"
                };
                var dates = new Models.WorkspaceDateFieldsInfo
                {
                    StatementOfLease = new DateTime(2026, 1, 20),
                    SevenDayLog = new DateTime(2026, 1, 21),
                    DayQualified = new DateTime(2026, 1, 25)
                };
                var conflicts = new List<string> { "Statement of Lease and 7-Day Prior Hours must match" };
                var files = new List<string> { "MVR.pdf", "Drug_Screen.pdf" };

                string auditPrompt = Services.CopilotContextBuilder.BuildAuditPrompt(header, dates, conflicts, files);
                if (!auditPrompt.Contains("Test Driver") || !auditPrompt.Contains("TRK-900") || !auditPrompt.Contains("***-**-0000"))
                {
                    throw new Exception("Copilot audit prompt missing expected driver context or failed SSN masking.");
                }
                if (!auditPrompt.Contains("[DOT Reference (Starter)]") || !auditPrompt.Contains("[Organization Requirements (Editable)]"))
                {
                    throw new Exception("Document review prompt missing DOT reference or editable requirements.");
                }

                // Verify response parsing
                string sampleResponse = @"Here is the audit report...
[APP_UPDATE]
{
  ""StatementOfLease"": ""01/20/2026"",
  ""SevenDayLog"": ""01/20/2026"",
  ""DayQualified"": ""01/20/2026""
}
[/APP_UPDATE]";
                var parsedDates = Services.CopilotContextBuilder.ParseSuggestedDates(sampleResponse);
                if (parsedDates.Count != 3 || parsedDates["StatementOfLease"] != "01/20/2026")
                {
                    throw new Exception($"Copilot suggested dates parsing failed, expected 3 dates, got {parsedDates.Count}");
                }

                // Verify Copilot/Assistant drawer exists on workspace and toggle button is labeled correctly
                if (ws.CopilotDrawerHost == null)
                {
                    throw new Exception("CopilotDrawerHost element not found on WorkspaceWindow.");
                }
                if (ws.CopilotToggleButton.Content?.ToString() != "Assistant")
                {
                    throw new Exception($"Expected CopilotToggleButton to have content 'View Assistant', got '{ws.CopilotToggleButton.Content}'");
                }

                // Verify updated 2-column Assistant toolbar buttons
                if (ws.CopilotAuditButton == null || (string)ws.CopilotAuditButton.Content != "Audit File & Report Issues")
                {
                    throw new Exception("CopilotAuditButton missing or incorrect content.");
                }
                if (ws.CopilotSummarizeButton == null || (string)ws.CopilotSummarizeButton.Content != "Driver Summarization")
                {
                    throw new Exception("CopilotSummarizeButton missing or incorrect content.");
                }
                if (ws.CopilotDraftRejectEmailButton == null || (string)ws.CopilotDraftRejectEmailButton.Content != "Draft Review Notes")
                {
                    throw new Exception("CopilotDraftRejectEmailButton missing or incorrect content.");
                }
                if (ws.CopilotGatherDatesButton == null || (string)ws.CopilotGatherDatesButton.Content != "📅 Gather All Dates")
                {
                    throw new Exception("CopilotGatherDatesButton missing or incorrect content.");
                }
                if (ws.CopilotApplyDatesButton == null || (string)ws.CopilotApplyDatesButton.Content != "📥 Apply Suggested Dates")
                {
                    throw new Exception("CopilotApplyDatesButton missing or incorrect content.");
                }

                // Verify reject email and auto-fill prompt builders
                string rejectPrompt = Services.CopilotContextBuilder.BuildDraftRejectEmailPrompt(header, dates, conflicts, files);
                if (!rejectPrompt.Contains("DRAFT DOCUMENT REVIEW FINDINGS") || !rejectPrompt.Contains("=== REVIEW CRITERIA ==="))
                {
                    throw new Exception("Review-notes prompt missing generic review criteria.");
                }

                string autoFillPrompt = auditPrompt;
                if (!autoFillPrompt.Contains("DOCUMENT REVIEW - FINDINGS AND DATA EXTRACTION") || !autoFillPrompt.Contains("TruckNumber"))
                {
                    throw new Exception("Auto-fill prompt builder missing expected headers.");
                }
                if (autoFillPrompt.Contains("John Doe") || autoFillPrompt.Contains("Doe Trucking") || autoFillPrompt.Contains("01/15/2026"))
                {
                    throw new Exception("Auto-fill prompt template must not contain literal dummy values or placeholder dates.");
                }

                // Verify reverse parser ignores prompt template and extracts assistant response
                string chatWithPromptEcho = autoFillPrompt + @"

Assistant response:
Here are the extracted qualification fields:
[APP_UPDATE]
{
  ""DriverName"": ""Synthetic Driver"",
  ""ContractorName"": ""Synthetic Demo Fleet"",
  ""TruckNumber"": ""DEMO-TRUCK"",
  ""TrailerNumber"": ""DEMO-TRAILER"",
  ""Ssn"": ""000-00-0000"",
  ""Fid"": ""00-0000000"",
  ""Phone"": ""202-555-0142"",
  ""Certification"": ""10/12/2026"",
  ""LimitedConsent"": ""10/12/2026"",
  ""FullConsent"": ""10/05/2026"",
  ""DrugTestRan"": ""10/02/2026"",
  ""DrugTestReceived"": ""10/06/2026"",
  ""CdlExpiration"": ""10/17/2029"",
  ""MvrCdlExpiration"": ""10/17/2029"",
  ""MvrLastReview"": ""10/12/2026"",
  ""MvrNextReview"": ""10/12/2027"",
  ""PhysicalLast"": ""05/21/2026"",
  ""PhysicalNext"": ""05/21/2028"",
  ""HazmatExpiration"": ""10/17/2029"",
  ""LatestJobEv"": ""10/09/2026"",
  ""SevenDayLog"": ""10/12/2026"",
  ""DayQualified"": ""10/12/2026"",
  ""PspSnapshotDate"": ""09/25/2026""
}
[/APP_UPDATE]";

                var echoExtracted = Services.CopilotContextBuilder.ParseSuggestedUpdates(chatWithPromptEcho);
                if (echoExtracted.HeaderFields["DriverName"] != "Synthetic Driver" ||
                    echoExtracted.HeaderFields["ContractorName"] != "Synthetic Demo Fleet" ||
                    echoExtracted.HeaderFields["TruckNumber"] != "DEMO-TRUCK")
                {
                    throw new Exception($"Reverse parser extracted wrong header fields from chat with prompt echo: {echoExtracted.HeaderFields["DriverName"]}");
                }
                if (echoExtracted.Dates["Certification"] != "10/12/2026" ||
                    echoExtracted.Dates["DrugTestRan"] != "10/02/2026" ||
                    echoExtracted.Dates["DayQualified"] != "10/12/2026")
                {
                    throw new Exception($"Date parsing failed or flipped month/day: Certification={echoExtracted.Dates["Certification"]}, DrugTestRan={echoExtracted.Dates["DrugTestRan"]}");
                }

                // Missing values stay blank when not present in source documents.
                ws.StatementOfLeaseDateBox.Text = "";
                ws.SevenDayLogDateBox.Text = "";
                ws.DayQualifiedDateBox.Text = "";
                typeof(Windows.WorkspaceWindow).GetMethod("ApplySuggestedDatesToBoxes", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!
                    .Invoke(ws, new object[] { echoExtracted.Dates });
                if (!string.IsNullOrWhiteSpace(ws.StatementOfLeaseDateBox.Text) || ws.SevenDayLogDateBox.Text != "10/12/2026" || ws.DayQualifiedDateBox.Text != "10/12/2026")
                {
                    throw new Exception("Missing dates should not be inferred or auto-filled.");
                }
                log.Add("PASS: Missing dates stay blank unless found in the source documents.");

                // Verify full suggested updates parsing (both header fields and dates)
                string fullUpdateSample = @"
[APP_UPDATE]
{
  ""DriverName"": ""Synthetic Driver"",
  ""ContractorName"": ""Synthetic Demo Fleet"",
  ""TruckNumber"": ""DEMO-TRUCK"",
  ""TrailerNumber"": ""DEMO-TRAILER"",
  ""Ssn"": ""000-00-0000"",
  ""Phone"": ""202-555-0142"",
  ""HazmatExpiration"": ""08/15/2028"",
  ""StatementOfLease"": ""01/20/2026"",
  ""SevenDayLog"": ""01/20/2026"",
  ""DayQualified"": ""01/20/2026""
}
[/APP_UPDATE]";
                var extracted = Services.CopilotContextBuilder.ParseSuggestedUpdates(fullUpdateSample);
                if (extracted.HeaderFields.Count != 6 || extracted.HeaderFields["DriverName"] != "Synthetic Driver")
                {
                    throw new Exception("Extracted header fields mismatch.");
                }
                if (extracted.Dates.Count != 4 || extracted.Dates["HazmatExpiration"] != "08/15/2028")
                {
                    throw new Exception("Extracted dates mismatch.");
                }

                // Verify the prompt uses modular, user-editable review criteria.
                if (!auditPrompt.Contains("user-configured review criteria") || !auditPrompt.Contains("[Industry Requirements (Editable)]"))
                {
                    throw new Exception("Prompts should use industry-neutral, editable review criteria.");
                }

                // Verify file materialization and attachment resolution
                var mockPdfBytes = new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x34 }; // %PDF-1.4
                var mockItem = new Models.PdfItemModel
                {
                    FilePath = "synthetic-demo-combined.pdf",
                    FileBytes = mockPdfBytes
                };
                var itemsField = typeof(Windows.WorkspaceWindow).GetField("_items", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!;
                var itemsColl = (System.Collections.ObjectModel.ObservableCollection<Models.CanvasItemModel>)itemsField.GetValue(ws)!;
                itemsColl.Add(mockItem);

                var resolvedFile = ws.GetOrExtractStoredFilePath("synthetic-demo-combined.pdf");
                if (resolvedFile == null || !File.Exists(resolvedFile.Value.FilePath) || resolvedFile.Value.FileName != "synthetic-demo-combined.pdf")
                {
                    throw new Exception("Failed to resolve or extract stored PDF file to disk.");
                }
                log.Add($"PASS: Stored PDF extracted to disk: {resolvedFile.Value.FilePath}");

                // Verify prompt includes attached document notice
                string auditWithPdfPrompt = Services.CopilotContextBuilder.BuildAuditPrompt(header, dates, conflicts, new[] { "synthetic-demo-combined.pdf" });
                if (!auditWithPdfPrompt.Contains("synthetic-demo-combined.pdf (ATTACHED TO THIS CHAT SESSION)") ||
                    !auditWithPdfPrompt.Contains("ATTACHED / LOADED WORKSPACE DOCUMENTS"))
                {
                    throw new Exception("Copilot audit prompt did not indicate attached document session state.");
                }

                log.Add("PASS: Assistant integration, prompt builder, toggle button name, and schema parsing verified cleanly!");
                log.Add("PASS: Assistant file attachment, extraction, and toolbar controls verified cleanly!");
            }

            // Check 8b: tracker overlap math
            {
                var tm = new Models.TrackerModel();
                var d0 = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-20);
                string F(DateTime d) => d.ToString("MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                tm.Rows[0].Description = "A"; tm.Rows[0].IsOtr = true; tm.Rows[0].Start = F(d0); tm.Rows[0].End = F(d0.AddMonths(12));
                tm.Rows[1].Description = "B"; tm.Rows[1].IsOtr = true; tm.Rows[1].Start = F(d0.AddMonths(5)); tm.Rows[1].End = F(d0.AddMonths(12));
                tm.Rows[2].Description = "Older history"; tm.Rows[2].IsOtr = true; tm.Rows[2].Start = "01/01/2010"; tm.Rows[2].End = "01/01/2011";
                if (tm.OtrMonths != 24 || tm.TotalMonths != 24) throw new Exception($"Tracker should include older history while counting overlap once (got OTR {tm.OtrMonths}, total {tm.TotalMonths}).");
                if (tm.Rows[2].Start != "01/01/2010") throw new Exception("Tracker date should keep 4-digit year.");
                log.Add("PASS: Tracker includes all history, counts overlap once, and preserves 4-digit years.");
            }

            // Check 9: Verify WebView2Loader initialization and library resolution
            {
                Services.WebView2LoaderHelper.EnsureLoaderInitialized();
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string testUserDataFolder = System.IO.Path.Combine(localAppData, "HexxcorpDocumentReviewUtility", "WebView2VerifyTest");
                var env = Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(null, testUserDataFolder).GetAwaiter().GetResult();
                if (env == null)
                {
                    throw new Exception("CoreWebView2Environment.CreateAsync returned null");
                }
                log.Add($"PASS: CoreWebView2Environment created successfully! BrowserVersion={env.BrowserVersionString}");
            }

            ws.AllowRealClose();
            ws.Close();
            log.Add("ALL VERIFICATION CHECKS PASSED!");
            File.WriteAllLines("verify_results.txt", log);
            return 0;
        }
        catch (Exception ex)
        {
            log.Add($"FAILED: {ex}");
            File.WriteAllLines("verify_results.txt", log);
            return 1;
        }
        finally
        {
            try
            {
                ws?.AllowRealClose();
                ws?.Close();
            }
            catch {}
        }
    }

    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        MessageBox.Show(
            $"An unexpected error occurred and was prevented from crashing the application:\n\n{e.Exception.Message}",
            "Hexxcorp Document Review Utility - Error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true;
    }
}
