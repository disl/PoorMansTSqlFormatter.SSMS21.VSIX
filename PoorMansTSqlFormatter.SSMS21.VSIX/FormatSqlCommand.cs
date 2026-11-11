using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using PoorMansTSqlFormatterLib;
using PoorMansTSqlFormatterLib.Formatters;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace PoorMansTSqlFormatter.SSMS21.VSIX
{
    /// <summary>
    /// Command handler: formatiert T-SQL mit Poor Man's T-SQL Formatter, wobei Kommentare
    /// erhalten bleiben. GUIDs/IDs passen zu FormatSqlCommandPackage.vsct.
    /// </summary>
    internal sealed class FormatSqlCommand
    {
        /// <summary>Command-ID „Auswahl formatieren" (matches „FormatSqlCommandId" im .vsct).</summary>
        public const int FormatSelectionCommandId = 0x0100;

        /// <summary>Command-ID „Ganzes Dokument formatieren" (matches „FormatWholeDocumentCommandId").</summary>
        public const int FormatWholeDocumentCommandId = 0x0101;

        /// <summary>Command-ID „Formatierungseinstellungen…" (matches „FormattingSettingsCommandId").</summary>
        public const int FormattingSettingsCommandId = 0x0102;

        /// <summary>Command-ID „Export in Excel" (matches „ExportToExcelCommandId").</summary>
        public const int ExportToExcelCommandId = 0x0103;

        /// <summary>Command-Set-GUID (matches „guidFormatSqlCommandPackageCmdSet" im .vsct).</summary>
        public static readonly Guid CommandSet = new Guid("3f53685b-f1e9-45b3-bc46-8bb825ab60a2");

        private readonly AsyncPackage package;

        private FormatSqlCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            if (commandService == null) throw new ArgumentNullException(nameof(commandService));

            AddCommand(commandService, FormatSelectionCommandId, this.ExecuteFormatSelection, this.QueryStatusWritableDocument);
            AddCommand(commandService, FormatWholeDocumentCommandId, this.ExecuteFormatWholeDocument, this.QueryStatusWritableDocument);
            AddCommand(commandService, FormattingSettingsCommandId, this.ExecuteSettings, this.QueryStatusLocalizeOnly);
            AddCommand(commandService, ExportToExcelCommandId, this.ExecuteExportToExcel, this.QueryStatusAnyDocument);
        }

        /// <summary>
        /// Initialisiert den Singleton der Kommandos.
        /// </summary>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // AddCommand muss auf dem UI-Thread laufen.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService =
                await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            var command = new FormatSqlCommand(package, commandService);
            command.EnsureKeyBindings();
            command.ShowToolbar();
        }

        /// <summary>
        /// SSMS-22-Besonderheit: KeyBindings aus dem .vsct werden nur in das VS-"Default"-Schema
        /// gemerged, nicht in das aktive SSMS-Schema ("(Standard)"). Deshalb werden die Shortcuts
        /// hier beim Package-Load direkt in das aktive Schema geschrieben. Idempotent: Es wird nur
        /// gesetzt, wenn das Kommando aktuell KEINE effektive Tastenzuordnung hat, damit spaetere
        /// Anpassungen des Nutzers (Extras &gt; Optionen &gt; Tastatur) nicht ueberschrieben werden.
        /// Die drei Kombos wurden vorab gegen die live SSMS-22-Belegung geprueft und sind dort frei.
        /// </summary>
        private void EnsureKeyBindings()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
                if (dte == null) return;

                var combos = new[]
                {
                    new { Id = FormatSelectionCommandId, Combo = "Global::Ctrl+Alt+F" },
                    new { Id = FormatWholeDocumentCommandId, Combo = "Global::Ctrl+Alt+D" },
                    new { Id = FormattingSettingsCommandId, Combo = "Global::Ctrl+Alt+K" },
                };

                foreach (var entry in combos)
                {
                    var cmd = FindCommand(dte, entry.Id);
                    if (cmd == null || HasEffectiveBinding(cmd)) continue;
                    cmd.Bindings = new object[] { entry.Combo };
                    Log("Tastenkombination gesetzt: " + entry.Combo);
                }

                EnsureExcelKeyBinding(dte);
            }
            catch (Exception ex)
            {
                Log("EnsureKeyBindings Fehler: " + ex.Message);
            }
        }

        /// <summary>
        /// Strg+Alt+E für „Export in Excel". Anders als bei den drei Formatier-Kombos ist hier
        /// nicht vorab gegen die SSMS-Belegung geprüft – deshalb wird nur gesetzt, wenn kein
        /// anderes Kommando die Kombination bereits nutzt. Die (teure) Prüfung läuft einmalig;
        /// danach verhindert eine Markierungsdatei weitere Versuche.
        /// </summary>
        private static void EnsureExcelKeyBinding(DTE2 dte)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            const string combo = "Global::Ctrl+Alt+E";
            string marker = Path.Combine(Path.GetDirectoryName(FormattingSettings.ConfigFilePath), "excel-shortcut-checked.txt");
            if (File.Exists(marker)) return;

            var cmd = FindCommand(dte, ExportToExcelCommandId);
            if (cmd == null) return; // Kommando noch nicht registriert -> naechster Start
            if (!HasEffectiveBinding(cmd))
            {
                string usedBy = FindCommandUsingKeys(dte, "Ctrl+Alt+E");
                if (usedBy == null)
                {
                    cmd.Bindings = new object[] { combo };
                    Log("Tastenkombination gesetzt: " + combo);
                }
                else
                {
                    Log("Strg+Alt+E nicht gesetzt, bereits belegt durch: " + usedBy);
                }
            }

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(marker));
                File.WriteAllText(marker, DateTime.Now.ToString("s"));
            }
            catch
            {
                // Dann wird beim naechsten Start erneut geprueft.
            }
        }

        /// <summary>Name des ersten Kommandos, das die Tastenkombination (in beliebigem Bereich) nutzt.</summary>
        private static string FindCommandUsingKeys(DTE2 dte, string keys)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            foreach (Command c in dte.Commands)
            {
                try
                {
                    if (!(c.Bindings is object[] bindings)) continue;
                    foreach (object b in bindings)
                    {
                        string s = b as string;
                        if (s == null) continue;
                        int idx = s.IndexOf("::", StringComparison.Ordinal);
                        string key = idx >= 0 ? s.Substring(idx + 2) : s;
                        if (string.Equals(key.Trim(), keys, StringComparison.OrdinalIgnoreCase))
                            return string.IsNullOrEmpty(c.Name) ? c.Guid + ":" + c.ID : c.Name;
                    }
                }
                catch
                {
                    // Einzelne Kommandos ohne lesbare Bindings ueberspringen.
                }
            }
            return null;
        }

        /// <summary>
        /// Findet unser Kommando anhand der GUID/ID in der DTE-Kommando-Collection.
        /// </summary>
        private static Command FindCommand(DTE2 dte, int commandId)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                string want = CommandSet.ToString("N").ToLowerInvariant();
                foreach (Command cmd in dte.Commands)
                {
                    string guid = (cmd.Guid ?? "").Replace("{", "").Replace("}", "").Replace("-", "").ToLowerInvariant();
                    if (guid.IndexOf(want, StringComparison.Ordinal) >= 0 && cmd.ID == commandId)
                        return cmd;
                }
            }
            catch
            {
                // Shell-Kommando-Collection nicht lesbar -> Shortcuts werden in diesem Lauf uebersprungen.
            }
            return null;
        }

        /// <summary>
        /// True, wenn das Kommando mindestens eine effektive Tastenzuordnung hat (leere
        /// "Global::"-Platzhalter zaehlen nicht). So bleiben Nutzer-Anpassungen erhalten.
        /// </summary>
        private static bool HasEffectiveBinding(Command cmd)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var bindings = cmd.Bindings as object[];
                if (bindings == null) return false;
                foreach (var b in bindings)
                {
                    string s = b as string;
                    if (string.IsNullOrWhiteSpace(s)) continue;
                    int idx = s.IndexOf("::", StringComparison.Ordinal);
                    string key = idx >= 0 ? s.Substring(idx + 2) : s;
                    if (!string.IsNullOrWhiteSpace(key)) return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private static void AddCommand(OleMenuCommandService service, int commandId,
            EventHandler executeHandler, EventHandler queryStatusHandler)
        {
            var menuCommandID = new CommandID(CommandSet, commandId);
            var menuItem = new OleMenuCommand(executeHandler, menuCommandID);
            if (queryStatusHandler != null)
                menuItem.BeforeQueryStatus += queryStatusHandler;
            service.AddCommand(menuItem);
        }

        /// <summary>
        /// Aktiviert ein Kommando nur, wenn ein beschreibbares Dokument aktiv ist (z.B. SQL-Editor).
        /// Beim Status-Abfragen werden zugleich die lokalen Captions angewandt (deutsches SSMS),
        /// damit auch das Untermenue korrekt beschriftet ist, wenn es angezeigt wird.
        /// </summary>
        private void QueryStatusWritableDocument(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ApplyLocalizedText(sender);
            if (sender is OleMenuCommand command)
            {
                var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
                command.Enabled = dte?.ActiveDocument != null && !dte.ActiveDocument.ReadOnly;
            }
        }

        /// <summary>
        /// „Export in Excel": aktiv, sobald ein Dokument offen ist. Ob es Raster-Ergebnisse gibt,
        /// wird erst beim Klick geprueft (Hinweis-Dialog), damit QueryStatus billig bleibt.
        /// </summary>
        private void QueryStatusAnyDocument(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ApplyLocalizedText(sender);
            if (sender is OleMenuCommand command)
            {
                var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
                command.Enabled = dte?.ActiveDocument != null;
            }
        }

        /// <summary>Exportiert die Ergebnisraster des aktiven Abfragefensters nach Excel.</summary>
        private void ExecuteExportToExcel(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ExcelExportCommand.Execute();
        }

        /// <summary>
        /// Nur-Lokalisierungs-QueryStatus fuer den Einstellungs-Button (immer aktiv).
        /// </summary>
        private void QueryStatusLocalizeOnly(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ApplyLocalizedText(sender);
        }

        /// <summary>
        /// Setzt die lokalisierte Beschriftung (deutsches SSMS) direkt am Kommando. Wirkt in
        /// Menue und Toolbar-Tooltip, weil die Buttons im .vsct das Flag TextChanges tragen.
        ///
        /// 11.11.26 ds: Ersetzt die fruehere Umbeschriftung ueber DTE.CommandBars
        /// (ApplyLocalizedCaptions/RelabelBar/RelabelControl). Die lief bei JEDER Statusabfrage
        /// rekursiv durch alle Menues ohne Zyklus-Erkennung; SSMS 22.12 hat einen zyklischen
        /// Menueeintrag (Erstellen &gt; Build enthaelt sich selbst) -> Endlosrekursion ->
        /// StackOverflow (0xC00000FD) -> SSMS-Absturz kurz nach dem Formatieren.
        /// </summary>
        private static void ApplyLocalizedText(object sender)
        {
            if (sender is OleMenuCommand command)
            {
                string text = Localizer.CommandText(command.CommandID.ID);
                if (text != null && command.Text != text)
                    command.Text = text;
            }
        }

        /// <summary>
        /// Formatiert die aktuelle Auswahl.
        /// </summary>
        private void ExecuteFormatSelection(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
            if (dte?.ActiveDocument?.Selection is TextSelection selection && !selection.IsEmpty)
            {
                EditPoint top = selection.TopPoint.CreateEditPoint();
                EditPoint bottom = selection.BottomPoint.CreateEditPoint();
                FormatAndReplace(top, bottom, top.GetText(bottom));
            }
            else
            {
                MessageBox.Show(Localizer.NoSelectionMessage,
                    "Poor Man's T-SQL Formatter", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Formatiert das gesamte aktive Dokument.
        /// </summary>
        private void ExecuteFormatWholeDocument(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
            if (dte?.ActiveDocument is Document doc)
            {
                // Das ganze Dokument über die TextDocument-Schnittstelle lesen.
                if (doc.Object("TextDocument") is TextDocument textDoc)
                {
                    EditPoint top = textDoc.StartPoint.CreateEditPoint();
                    EditPoint bottom = textDoc.EndPoint.CreateEditPoint();
                    FormatAndReplace(top, bottom, top.GetText(bottom));
                }
            }
        }

        /// <summary>
        /// Macht die Symbolleiste "Poor Man's T-SQL Formatter" sichtbar. Toolbars aus dem
        /// .vsct sind standardmaessig unsichtbar (nur ueber Ansicht &gt; Symbolleisten
        /// aktivierbar); damit die Leiste direkt nach der Installation unten an der
        /// Menueleiste angedockt erscheint, setzen wir hier die Visible-Eigenschaft.
        /// CommandBars wird bewusst per Late-Binding angesprochen: Die Interop-Typen
        /// (EnvDTE.CommandBars) liegen in SSMS 22 in einer leeren Facade-Assembly und
        /// sind zur Compile-Zeit nicht aufloesbar; ueber Reflection funktioniert der
        /// COM-RCW zur Laufzeit zuverlaessig.
        /// </summary>
        private void ShowToolbar()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
                if (dte == null) return;
                object commandBars = dte.CommandBars;
                if (commandBars == null) return;

                object bar = commandBars.GetType().InvokeMember(
                    "Item", System.Reflection.BindingFlags.GetProperty, null, commandBars,
                    new object[] { "Poor Man's T-SQL Formatter" });
                if (bar == null) return;

                bool visible = (bool)bar.GetType().InvokeMember(
                    "Visible", System.Reflection.BindingFlags.GetProperty, null, bar, null);
                if (!visible)
                    bar.GetType().InvokeMember(
                        "Visible", System.Reflection.BindingFlags.SetProperty, null, bar,
                        new object[] { true });
            }
            catch (Exception ex)
            {
                Log("ShowToolbar Fehler: " + ex.Message);
            }
        }

        /// <summary>
        /// Öffnet den Einstellungs-Dialog und speichert die Optionen.
        /// </summary>
        private void ExecuteSettings(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            using (var form = new FormattingSettingsForm(FormattingSettings.Current))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    FormattingSettings.Save(form.Result);
                    Log("Einstellungen gespeichert: " + form.Result.ToSerializedString());
                }
            }
        }

        /// <summary>
        /// Kern: Text formatieren, Fehler-Dialog, dann den Bereich von startPoint bis endPoint
        /// ersetzen. Die Auswahl wird danach auf einen leeren Cursor am Anfang zusammengezogen.
        /// </summary>
        private void FormatAndReplace(EditPoint startPoint, EditPoint endPoint, string input)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool errorsFound = false;
            string formatted;
            var sw = new Stopwatch();
            sw.Start();
            try
            {
                var formatter = new TSqlStandardFormatter(FormattingSettings.Current);
                formatted = new SqlFormattingManager(formatter).Format(input, ref errorsFound);
            }
            catch (Exception ex)
            {
                sw.Stop();
                Log(string.Format("Formatter-Ausnahme nach {0} ms: {1}", sw.ElapsedMilliseconds, ex.Message));
                MessageBox.Show(Localizer.FormatFailedPrefix + ex.Message,
                    "Poor Man's T-SQL Formatter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            sw.Stop();
            Log(string.Format("Format: {0} ms, Errors={1}, {2} -> {3} Zeichen",
                sw.ElapsedMilliseconds, errorsFound, input.Length, formatted.Length));

            if (errorsFound)
            {
                DialogResult answer = MessageBox.Show(
                    Localizer.ParseWarningMessage,
                    "Poor Man's T-SQL Formatter", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.Yes)
                {
                    Log("Abgebrochen (Errors=true)");
                    return;
                }
            }

            sw.Restart();
            startPoint.ReplaceText(endPoint, formatted,
                (int)vsEPReplaceTextOptions.vsEPReplaceTextKeepMarkers);
            sw.Stop();
            Log(string.Format("ReplaceText: {0} ms", sw.ElapsedMilliseconds));

            // Auswahl auf einen leeren Cursor am Anfang des formatierten Texts zusammenziehen.
            sw.Restart();
            var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
            (dte?.ActiveDocument?.Selection as TextSelection)?.MoveToPoint(startPoint, false);
            sw.Stop();
            Log(string.Format("MoveToPoint: {0} ms", sw.ElapsedMilliseconds));
        }

        /// <summary>
        /// Hängt eine kurze Diagnose-Zeile an die Logdatei (UTF-8) an. Fehler beim Schreiben
        /// werden geschluckt, damit das Logging nie den Format-Vorgang stört.
        /// </summary>
        internal static void Log(string message)
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "poor_mans_formatter_log.txt");
                File.AppendAllText(path,
                    string.Format("[{0:HH:mm:ss.fff}] {1}{2}", DateTime.Now, message, Environment.NewLine),
                    new System.Text.UTF8Encoding(false));
            }
            catch
            {
                // Logging nie kritisch machen.
            }
        }
    }
}
