using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Threading;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PoorMansTSqlFormatter.SSMS21.VSIX
{
    /// <summary>
    /// Button „Export in Excel": exportiert alle Ergebnisraster des aktiven Abfragefensters
    /// typgerecht in eine .xlsx-Datei und öffnet sie (Standardprogramm für .xlsx, i.d.R. Excel).
    /// Die Abfrage wird NICHT erneut ausgeführt – exportiert wird, was im Raster steht.
    /// </summary>
    internal static class ExcelExportCommand
    {
        private const string Caption = "Poor Man's T-SQL Formatter – Excel";

        public static void Execute()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            List<ResultSetSnapshot> resultSets;
            try
            {
                resultSets = ResultGridReader.ReadActiveWindow();
            }
            catch (Exception ex)
            {
                ResultGridReader.Log("Lesen der Ergebnisse fehlgeschlagen: " + ex);
                MessageBox.Show(Localizer.ExcelReadFailedPrefix + Unwrap(ex).Message, Caption,
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (resultSets.Count == 0)
            {
                MessageBox.Show(Localizer.ExcelNoResultsMessage, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            long totalRows = 0;
            foreach (ResultSetSnapshot rs in resultSets) totalRows += rs.RowCount;

            string path;
            try
            {
                string dir = Path.Combine(Path.GetTempPath(), "SSMS-Excel-Export");
                Directory.CreateDirectory(dir);
                path = Path.Combine(dir, "SSMS_Export_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(Localizer.ExcelExportFailedPrefix + ex.Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ExcelExportStats stats = null;
            Exception error = null;
            bool cancelled = false;
            var sw = Stopwatch.StartNew();

            // Export im Hintergrund; SSMS zeigt nach 1 s einen Fortschrittsdialog mit „Abbrechen".
            ThreadHelper.JoinableTaskFactory.Run(
                Localizer.ExcelWaitTitle,
                Localizer.ExcelWaitMessage,
                async (progress, ct) =>
                {
                    await TaskScheduler.Default;
                    var rowProgress = new Progress<long>(done =>
                        progress.Report(new ThreadedWaitDialogProgressData(
                            Localizer.ExcelWaitMessage,
                            Localizer.ExcelProgressText(done, totalRows),
                            null, true,
                            (int)Math.Min(done, int.MaxValue), (int)Math.Min(Math.Max(totalRows, 1), int.MaxValue))));
                    try
                    {
                        stats = ExcelExporter.Export(resultSets, path, rowProgress, ct);
                    }
                    catch (OperationCanceledException)
                    {
                        cancelled = true;
                    }
                    catch (Exception ex)
                    {
                        error = ex;
                    }
                },
                TimeSpan.FromSeconds(1));

            sw.Stop();

            if (cancelled || error != null)
            {
                TryDelete(path);
                if (error != null)
                {
                    ResultGridReader.Log("Fehler: " + error);
                    MessageBox.Show(Localizer.ExcelExportFailedPrefix + Unwrap(error).Message, Caption,
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    ResultGridReader.Log("abgebrochen");
                }
                return;
            }

            ResultGridReader.Log(string.Format(
                "{0} Zeilen, {1} Blätter, {2} ms, gekürzt={3}, Zahlen>15 Stellen als Text={4}, Datum<1900 als Text={5}, unvollständig={6} -> {7}",
                stats.Rows, stats.Sheets, sw.ElapsedMilliseconds, stats.TruncatedTexts,
                stats.PreciseNumbersAsText, stats.OldDatesAsText, stats.IncompleteResultSets, path));

            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                // Kein Programm für .xlsx registriert → Speicherort nennen.
                ResultGridReader.Log("Öffnen fehlgeschlagen: " + ex.Message);
                MessageBox.Show(Localizer.ExcelOpenFailedMessage(path), Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string notes = Localizer.ExcelNotes(stats);
            if (notes != null)
                MessageBox.Show(notes, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static Exception Unwrap(Exception ex)
        {
            while (ex is System.Reflection.TargetInvocationException && ex.InnerException != null)
                ex = ex.InnerException;
            return ex;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch
            {
                // Temp-Datei bleibt liegen – unkritisch.
            }
        }
    }
}
