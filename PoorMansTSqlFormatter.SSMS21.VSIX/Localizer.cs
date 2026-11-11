using System.Collections.Generic;
using System.Globalization;

namespace PoorMansTSqlFormatter.SSMS21.VSIX
{
    /// <summary>
    /// Zentrale Laufzeit-Lokalisierung: deutsches SSMS bekommt deutsche Texte, alle
    /// anderen Systeme englische. Die .vsct-Defaults sind Englisch; auf deutschen
    /// Systemen setzen die QueryStatus-Handler den Text am Kommando (siehe
    /// <see cref="CommandText"/>, Flag TextChanges im .vsct).
    /// </summary>
    internal static class Localizer
    {
        /// <summary>True, wenn das SSMS-System deutsch ist (de-DE, de-AT, de-CH, ...).</summary>
        public static bool IsGerman =>
            CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "de";

        /// <summary>Waehlt je nach Systemsprache den deutschen oder englischen Text.</summary>
        public static string T(string german, string english) => IsGerman ? german : english;

        /// <summary>
        /// Beschriftung eines Kommandos (Command-ID aus dem .vsct) in der Systemsprache.
        /// Wird in den QueryStatus-Handlern als OleMenuCommand.Text gesetzt.
        /// </summary>
        public static string CommandText(int commandId)
        {
            switch (commandId)
            {
                case FormatSqlCommand.FormatSelectionCommandId: return T("Auswahl formatieren", "Format Selection");
                case FormatSqlCommand.FormatWholeDocumentCommandId: return T("Ganzes Dokument formatieren", "Format Whole Document");
                case FormatSqlCommand.FormattingSettingsCommandId: return T("Formatierungseinstellungen…", "Formatting Settings…");
                case FormatSqlCommand.ExportToExcelCommandId: return T("Export in Excel", "Export to Excel");
                default: return null;
            }
        }

        // --- Excel-Export ---

        public static string ExcelNoResultsMessage =>
            T(
                "Im aktiven Abfragefenster gibt es keine Ergebnisse im Raster.\r\n\r\n" +
                "Bitte die Abfrage mit „Ergebnisse in Raster“ (Strg+D) ausführen und dann erneut exportieren.",
                "The active query window has no grid results.\r\n\r\n" +
                "Run the query with \"Results to Grid\" (Ctrl+D) and export again.");

        public static string ExcelReadFailedPrefix =>
            T("Die Ergebnisse konnten nicht gelesen werden: ", "Could not read the results: ");

        public static string ExcelExportFailedPrefix =>
            T("Export in Excel fehlgeschlagen: ", "Export to Excel failed: ");

        public static string ExcelWaitTitle =>
            T("Export in Excel", "Export to Excel");

        public static string ExcelWaitMessage =>
            T("Ergebnisse werden nach Excel exportiert …", "Exporting results to Excel …");

        public static string ExcelProgressText(long done, long total) =>
            T(string.Format(CultureInfo.CurrentCulture, "{0:N0} von {1:N0} Zeilen", done, total),
              string.Format(CultureInfo.CurrentCulture, "{0:N0} of {1:N0} rows", done, total));

        public static string ExcelOpenFailedMessage(string path) =>
            T("Die Datei wurde erstellt, konnte aber nicht geöffnet werden (kein Programm für .xlsx registriert?):\r\n\r\n" + path,
              "The file was created but could not be opened (no program registered for .xlsx?):\r\n\r\n" + path);

        /// <summary>Hinweise nach dem Export; null, wenn es nichts zu melden gibt.</summary>
        public static string ExcelNotes(ExcelExportStats s)
        {
            var lines = new List<string>();
            if (s.IncompleteResultSets > 0)
                lines.Add(T("• Die Abfrage lief noch – exportiert wurden nur die bis dahin geladenen Zeilen.",
                            "• The query was still running – only the rows loaded so far were exported."));
            if (s.PreciseNumbersAsText > 0)
                lines.Add(T(string.Format(CultureInfo.CurrentCulture,
                                "• {0:N0} Zahl(en) mit mehr als 15 Stellen wurden als Text exportiert, damit keine Stellen verloren gehen (Excel rechnet nur mit 15 Stellen).",
                                s.PreciseNumbersAsText),
                            string.Format(CultureInfo.CurrentCulture,
                                "• {0:N0} number(s) with more than 15 digits were exported as text so that no digits are lost (Excel only keeps 15 digits).",
                                s.PreciseNumbersAsText)));
            if (s.OldDatesAsText > 0)
                lines.Add(T(string.Format(CultureInfo.CurrentCulture,
                                "• {0:N0} Datumswert(e) vor dem 01.03.1900 wurden als Text exportiert (in Excel nicht darstellbar).",
                                s.OldDatesAsText),
                            string.Format(CultureInfo.CurrentCulture,
                                "• {0:N0} date value(s) before 1900-03-01 were exported as text (not representable in Excel).",
                                s.OldDatesAsText)));
            if (s.TruncatedTexts > 0)
                lines.Add(T(string.Format(CultureInfo.CurrentCulture,
                                "• {0:N0} Text(e) wurden auf 32.767 Zeichen gekürzt (Excel-Grenze je Zelle).",
                                s.TruncatedTexts),
                            string.Format(CultureInfo.CurrentCulture,
                                "• {0:N0} text value(s) were truncated to 32,767 characters (Excel cell limit).",
                                s.TruncatedTexts)));
            if (lines.Count == 0) return null;
            return T("Export abgeschlossen. Hinweise:", "Export finished. Notes:") + "\r\n\r\n" + string.Join("\r\n", lines);
        }

        // --- MessageBox-Texte ---

        /// <summary>Keine Auswahl markiert (Info-Dialog).</summary>
        public static string NoSelectionMessage =>
            T(
                "Es ist keine Textauswahl markiert.\r\n\r\n" +
                "Für das komplette Skript: Menüpunkt „Ganzes Dokument formatieren“.",
                "No text is selected.\r\n\r\n" +
                "For the whole script: use the menu item \"Format Whole Document\".");

        /// <summary>Prefix der Fehler-MessageBox (Formatieren fehlgeschlagen).</summary>
        public static string FormatFailedPrefix =>
            T("Formatieren fehlgeschlagen: ", "Formatting failed: ");

        /// <summary>Warnung bei nicht vollständig parsebarem Text (Ja/Nein-Dialog).</summary>
        public static string ParseWarningMessage =>
            T(
                "Die Auswahl konnte nicht vollständig geparst werden.\r\n\r\n" +
                "Tipp: Die vollständige Prozedur markieren (Strg+A oder von ALTER PROCEDURE bis END).\r\n\r\n" +
                "Bei „Ja“ kann das Ergebnis unbrauchbar werden (Treppen-Einrückung) und ersetzt dein SQL.",
                "The selection could not be parsed completely.\r\n\r\n" +
                "Tip: Select the whole procedure (Ctrl+A or from ALTER PROCEDURE to END).\r\n\r\n" +
                "If you choose \"Yes\", the result may be unusable (staircase indentation) and it will replace your SQL.");

        // --- Einstellungs-Dialog ---

        public static string SettingsTitle =>
            T("Poor Man's T-SQL Formatter – Einstellungen", "Poor Man's T-SQL Formatter – Settings");

        public static string[] IndentChoices =>
            new[] { T("Tabulator", "Tab"), T("2 Leerzeichen", "2 spaces"), T("4 Leerzeichen", "4 spaces") };

        public static string IndentGroup =>
            T("Einrückung", "Indentation");

        public static string IndentLabel =>
            T("Einrückung:", "Indentation:");

        public static string SpacesPerTabLabel =>
            T("SpacesPerTab (Breite je Tab):", "SpacesPerTab (width per tab):");

        public static string MaxLineWidthLabel =>
            T("MaxLineWidth (Umbruchgrenze):", "MaxLineWidth (wrap limit):");

        public static string BlankLinesGroup =>
            T("Leerzeilen", "Blank lines");

        public static string StatementBreaksLabel =>
            T("Zwischen Statements (NewStatementLineBreaks):", "Between statements (NewStatementLineBreaks):");

        public static string ClauseBreaksLabel =>
            T("Zwischen Klauseln (NewClauseLineBreaks):", "Between clauses (NewClauseLineBreaks):");

        public static string ExpandListsGroup =>
            T("Listen aufklappen", "Expand lists");

        public static string CommaGroup =>
            T("Kommas", "Commas");

        public static string KeywordGroup =>
            T("Keywords", "Keywords");

        public static string ExpandCommaListsCheck =>
            T("Komma-Listen auf mehrere Zeilen\r\n(ExpandCommaLists)", "Comma lists on multiple lines\r\n(ExpandCommaLists)");

        public static string ExpandBooleanExpressionsCheck =>
            T("AND/OR auf eigene Zeilen\r\n(ExpandBooleanExpressions)", "AND/OR on separate lines\r\n(ExpandBooleanExpressions)");

        public static string ExpandCaseStatementsCheck =>
            T("CASE/WHEN/THEN untereinander\r\n(ExpandCaseStatements)", "CASE/WHEN/THEN stacked\r\n(ExpandCaseStatements)");

        public static string ExpandInListsCheck =>
            T("IN-Listen untereinander\r\n(ExpandInLists)", "IN lists stacked\r\n(ExpandInLists)");

        public static string ExpandBetweenConditionsCheck =>
            T("BETWEEN auf eigene Zeilen\r\n(ExpandBetweenConditions)", "BETWEEN on separate lines\r\n(ExpandBetweenConditions)");

        public static string BreakJoinOnSectionsCheck =>
            T("JOIN…ON auf eigene Zeilen\r\n(BreakJoinOnSections)", "JOIN…ON on separate lines\r\n(BreakJoinOnSections)");

        public static string TrailingCommasCheck =>
            T(
                "Kommas ans Zeilenende (TrailingCommas)\r\n– deaktiviert = führende Kommas",
                "Commas at end of line (TrailingCommas)\r\n– disabled = leading commas");

        public static string SpaceAfterExpandedCommaCheck =>
            T(
                "Leerzeichen nach Komma in Listen (SpaceAfterExpandedComma)",
                "Space after comma in lists (SpaceAfterExpandedComma)");

        public static string UppercaseKeywordsCheck =>
            T("Keywords in Großbuchstaben (UppercaseKeywords)", "Keywords in uppercase (UppercaseKeywords)");

        public static string KeywordStandardizationCheck =>
            T("Keywords standardisieren (KeywordStandardization)", "Standardize keywords (KeywordStandardization)");

        public static string ResetButton =>
            T("Standard wiederherstellen", "Restore defaults");

        public static string CancelButton =>
            T("Abbrechen", "Cancel");
    }
}
