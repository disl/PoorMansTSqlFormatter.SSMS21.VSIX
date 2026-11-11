using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Globalization;
using System.Text;
using System.Threading;

namespace PoorMansTSqlFormatter.SSMS21.VSIX
{
    /// <summary>Statistik eines Exports (für Log und Hinweis-Dialog).</summary>
    internal sealed class ExcelExportStats
    {
        public long Rows;
        public int Sheets;

        /// <summary>Texte, die auf 32.767 Zeichen (Excel-Grenze) gekürzt wurden.</summary>
        public long TruncatedTexts;

        /// <summary>Zahlen mit mehr als 15 signifikanten Stellen, die als Text exportiert wurden.</summary>
        public long PreciseNumbersAsText;

        /// <summary>Datumswerte vor 01.03.1900 (in Excel nicht darstellbar), als Text exportiert.</summary>
        public long OldDatesAsText;

        /// <summary>Ergebnisse, deren Abfrage beim Export noch lief.</summary>
        public int IncompleteResultSets;
    }

    /// <summary>
    /// Wandelt die SSMS-Ergebnisraster in eine typisierte .xlsx-Datei um.
    ///
    /// Grundsatz: Es wird nicht der angezeigte Text exportiert, sondern der typisierte Wert,
    /// den SSMS intern speichert. Zahlen und Datumswerte werden kulturunabhängig (invariant)
    /// in die Datei geschrieben; Excel zeigt sie anschließend in der Ländereinstellung des
    /// Anwenders an (deutsch: 1.234,56 / 31.12.2026). Dadurch entfallen die bekannten
    /// Probleme beim Kopieren/CSV-Import (Dezimalpunkt vs. Komma, Datum als Text, Zahlen
    /// als Datum, verlorene führende Nullen, wissenschaftliche Notation bei langen Nummern).
    ///
    /// Typ-Abbildung:
    ///   tinyint/smallint/int, bigint ≤ 15 Stellen           → Zahl, Format "0"
    ///   bigint/decimal/numeric > 15 signifikante Stellen     → Text (Excel rechnet nur mit 15 Stellen)
    ///   decimal/numeric(p,s)                                 → Zahl, s Nachkommastellen
    ///   money/smallmoney                                     → Zahl, "#,##0.00##"
    ///   float/real                                           → Zahl, Standardformat (Roundtrip-genau)
    ///   bit                                                  → Zahl 0/1 (wie im SSMS-Raster)
    ///   date                                                 → Datum
    ///   datetime/datetime2/smalldatetime                     → Datum + Uhrzeit (ms, falls vorhanden)
    ///   time                                                 → Uhrzeit
    ///   datetimeoffset                                       → Text inkl. Zeitzone (Excel kennt keine Zeitzonen)
    ///   uniqueidentifier                                     → Text (Großbuchstaben, wie SSMS)
    ///   binary/varbinary/image/timestamp                     → Text "0x…"
    ///   char/varchar/nchar/nvarchar/text/xml/sonstige        → Text (nie als Zahl/Datum/Formel interpretiert)
    ///   NULL                                                 → leere Zelle
    /// </summary>
    internal sealed class ExcelExporter
    {
        /// <summary>Excel speichert Zahlen als double: nur 15 signifikante Stellen sind exakt.</summary>
        private const int MaxExactDigits = 15;

        private static readonly DateTime ExcelMinDate = new DateTime(1900, 3, 1);
        private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        private readonly XlsxWriter writer;
        private readonly ExcelExportStats stats = new ExcelExportStats();
        private readonly int styleInteger;
        private readonly int styleMoney;
        private readonly int styleDate;
        private readonly int styleDateTime;
        private readonly int styleDateTimeMs;
        private readonly int styleSmallDateTime;
        private readonly int styleTime;
        private readonly int styleTimeMs;
        private readonly Dictionary<int, int> decimalStyles = new Dictionary<int, int>();

        private ExcelExporter(XlsxWriter writer)
        {
            this.writer = writer;
            string date = ExcelDatePattern(CultureInfo.CurrentCulture);
            styleInteger = writer.GetStyle("0");
            styleMoney = writer.GetStyle("#,##0.00##");
            styleDate = writer.GetStyle(date);
            styleDateTime = writer.GetStyle(date + " hh:mm:ss");
            styleDateTimeMs = writer.GetStyle(date + " hh:mm:ss.000");
            styleSmallDateTime = writer.GetStyle(date + " hh:mm");
            styleTime = writer.GetStyle("hh:mm:ss");
            styleTimeMs = writer.GetStyle("hh:mm:ss.000");
        }

        /// <summary>
        /// Exportiert alle Ergebnisraster in eine .xlsx-Datei (je Raster ein Arbeitsblatt;
        /// mehr als 1.048.575 Zeilen werden auf Folgeblätter verteilt). Läuft im Hintergrund.
        /// </summary>
        public static ExcelExportStats Export(IList<ResultSetSnapshot> resultSets, string path,
            IProgress<long> rowsDone, CancellationToken ct)
        {
            using (var writer = new XlsxWriter(path))
            {
                var exporter = new ExcelExporter(writer);
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < resultSets.Count; i++)
                {
                    string baseName = resultSets.Count == 1 && resultSets[i].RowCount < XlsxWriter.MaxRowsPerSheet
                        ? Localizer.T("Ergebnis", "Result")
                        : Localizer.T("Ergebnis ", "Result ") + (i + 1).ToString(Inv);
                    exporter.ExportResultSet(resultSets[i], baseName, usedNames, rowsDone, ct);
                }
                return exporter.stats;
            }
        }

        private void ExportResultSet(ResultSetSnapshot rs, string baseName, HashSet<string> usedNames,
            IProgress<long> rowsDone, CancellationToken ct)
        {
            if (!rs.Complete) stats.IncompleteResultSets++;
            ResultColumn[] cols = rs.Columns;
            double[] widths = EstimateWidths(rs);

            var header = new XlsxCell[cols.Length];
            for (int c = 0; c < cols.Length; c++)
                header[c] = XlsxCell.Text(XlsxWriter.SanitizeText(cols[c].Name), XlsxWriter.HeaderStyle);

            var cells = new XlsxCell[cols.Length];
            long row = 0;
            int part = 1;
            do
            {
                string name = UniqueSheetName(part == 1 ? baseName : baseName + " (" + part.ToString(Inv) + ")", usedNames);
                using (XlsxWriter.SheetWriter sheet = writer.BeginSheet(name, widths))
                {
                    stats.Sheets++;
                    sheet.WriteRow(header);
                    long rowsInSheet = 1;
                    while (row < rs.RowCount && rowsInSheet < XlsxWriter.MaxRowsPerSheet)
                    {
                        for (int c = 0; c < cols.Length; c++)
                            cells[c] = Convert(rs.GetValue(row, c), cols[c]);
                        sheet.WriteRow(cells);
                        row++;
                        rowsInSheet++;
                        stats.Rows++;
                        if ((stats.Rows & 0x3FF) == 0)
                        {
                            ct.ThrowIfCancellationRequested();
                            rowsDone?.Report(stats.Rows);
                        }
                    }
                }
                part++;
            }
            while (row < rs.RowCount);
            rowsDone?.Report(stats.Rows);
        }

        /// <summary>Spaltenbreiten aus Kopfzeile und den ersten 200 Zeilen schätzen.</summary>
        private double[] EstimateWidths(ResultSetSnapshot rs)
        {
            var widths = new double[rs.Columns.Length];
            long sample = Math.Min(rs.RowCount, 200);
            for (int c = 0; c < widths.Length; c++)
            {
                int max = (rs.Columns[c].Name ?? "").Length + 3; // + Platz für den Filter-Pfeil
                string type = rs.Columns[c].SqlType;
                if (type == "date") max = Math.Max(max, 11);
                else if (type.StartsWith("datetime", StringComparison.Ordinal) || type == "smalldatetime") max = Math.Max(max, 20);
                for (long r = 0; r < sample; r++)
                {
                    object v = rs.GetValue(r, c);
                    if (v == null || v is DBNull || (v is INullable n && n.IsNull)) continue;
                    string s = v is byte[] || v is SqlBinary ? "0x" : System.Convert.ToString(v, Inv) ?? "";
                    int nl = s.IndexOf('\n');
                    int len = nl >= 0 ? nl : s.Length;
                    if (len > max) max = len;
                    if (max >= 60) break;
                }
                widths[c] = Math.Min(Math.Max(max + 1, 6), 60);
            }
            return widths;
        }

        /// <summary>Wandelt einen von SSMS gespeicherten Zellwert in eine typisierte Excel-Zelle.</summary>
        private XlsxCell Convert(object value, ResultColumn col)
        {
            if (value == null || value is DBNull) return XlsxCell.Empty;
            if (value is INullable nullable && nullable.IsNull) return XlsxCell.Empty;

            switch (value)
            {
                // --- Ganzzahlen ---
                case SqlBoolean b: return XlsxCell.Number(b.Value ? "1" : "0", styleInteger);
                case bool b: return XlsxCell.Number(b ? "1" : "0", styleInteger);
                case SqlByte v: return XlsxCell.Number(v.Value.ToString(Inv), styleInteger);
                case SqlInt16 v: return XlsxCell.Number(v.Value.ToString(Inv), styleInteger);
                case SqlInt32 v: return XlsxCell.Number(v.Value.ToString(Inv), styleInteger);
                case SqlInt64 v: return Integer(v.Value);
                case byte v: return XlsxCell.Number(v.ToString(Inv), styleInteger);
                case sbyte v: return XlsxCell.Number(v.ToString(Inv), styleInteger);
                case short v: return XlsxCell.Number(v.ToString(Inv), styleInteger);
                case ushort v: return XlsxCell.Number(v.ToString(Inv), styleInteger);
                case int v: return XlsxCell.Number(v.ToString(Inv), styleInteger);
                case uint v: return XlsxCell.Number(v.ToString(Inv), styleInteger);
                case long v: return Integer(v);

                // --- Dezimal / Geld ---
                case SqlDecimal v: return Decimal(v, col);
                case decimal v: return Decimal(new SqlDecimal(v), col);
                case SqlMoney v: return Money(v.Value);

                // --- Gleitkomma ---
                case SqlDouble v: return Double(v.Value);
                case double v: return Double(v);
                case SqlSingle v: return Single(v.Value);
                case float v: return Single(v);

                // --- Datum / Zeit ---
                case SqlDateTime v: return DateTimeCell(v.Value, col);
                case DateTime v: return DateTimeCell(v, col);
                case TimeSpan v: return TimeCell(v, col);
                case DateTimeOffset v: return XlsxCell.Text(FormatDateTimeOffset(v, col));

                // --- Sonstige ---
                case SqlGuid v: return XlsxCell.Text(v.Value.ToString("D").ToUpperInvariant());
                case Guid v: return XlsxCell.Text(v.ToString("D").ToUpperInvariant());
                case SqlBinary v: return Text(HexString(v.Value));
                case byte[] v: return Text(HexString(v));
                case SqlBytes v: return Text(HexString(v.Value));
                case SqlString v: return StringValue(v.Value, col);
                case SqlChars v: return StringValue(new string(v.Value), col);
                case SqlXml v: return Text(v.Value);
                case string v: return StringValue(v, col);
                default:
                    return Text(System.Convert.ToString(value, Inv));
            }
        }

        private XlsxCell Integer(long v)
        {
            if (v > 999999999999999L || v < -999999999999999L)
            {
                stats.PreciseNumbersAsText++;
                return XlsxCell.Text(v.ToString(Inv));
            }
            return XlsxCell.Number(v.ToString(Inv), styleInteger);
        }

        private XlsxCell Decimal(SqlDecimal v, ResultColumn col)
        {
            // SqlDecimal.ToString() ist kulturunabhängig und verlustfrei (bis 38 Stellen).
            string s = v.ToString();
            if (SignificantDigits(s) > MaxExactDigits)
            {
                stats.PreciseNumbersAsText++;
                return XlsxCell.Text(s);
            }
            int scale = col.Scale >= 0 && col.Scale <= 38 ? col.Scale : v.Scale;
            if (col.SqlType == "money" || col.SqlType == "smallmoney") return XlsxCell.Number(s, styleMoney);
            return XlsxCell.Number(s, DecimalStyle(scale));
        }

        private XlsxCell Money(decimal v)
        {
            string s = v.ToString(Inv);
            if (SignificantDigits(s) > MaxExactDigits)
            {
                stats.PreciseNumbersAsText++;
                return XlsxCell.Text(s);
            }
            return XlsxCell.Number(s, styleMoney);
        }

        private static XlsxCell Double(double v)
        {
            if (double.IsNaN(v) || double.IsInfinity(v)) return XlsxCell.Text(v.ToString(Inv));
            return XlsxCell.Number(v.ToString("R", Inv), 0);
        }

        private static XlsxCell Single(float v)
        {
            if (float.IsNaN(v) || float.IsInfinity(v)) return XlsxCell.Text(v.ToString(Inv));
            // Über den float-Roundtrip-String, sonst wird 0.1f zu 0.100000001490116.
            return XlsxCell.Number(v.ToString("R", Inv), 0);
        }

        private int DecimalStyle(int scale)
        {
            if (scale > 15) scale = 15;
            if (!decimalStyles.TryGetValue(scale, out int style))
            {
                style = scale <= 0 ? styleInteger : writer.GetStyle("#,##0." + new string('0', scale));
                decimalStyles[scale] = style;
            }
            return style;
        }

        private XlsxCell DateTimeCell(DateTime v, ResultColumn col)
        {
            string type = col.SqlType;
            if (v < ExcelMinDate)
            {
                // Excel kann Daten vor 1900 nicht darstellen (und 1900 hat einen Schaltjahr-Fehler).
                stats.OldDatesAsText++;
                return XlsxCell.Text(type == "date" ? v.ToString("yyyy-MM-dd", Inv) : v.ToString("yyyy-MM-dd HH:mm:ss.fff", Inv));
            }

            string serial = v.ToOADate().ToString("R", Inv);
            int style;
            if (type == "date") style = styleDate;
            else if (type == "smalldatetime") style = styleSmallDateTime;
            else if (v.Millisecond != 0 || (v.Ticks % TimeSpan.TicksPerMillisecond) != 0) style = styleDateTimeMs;
            else if (type.Length == 0 && v.TimeOfDay == TimeSpan.Zero) style = styleDate;
            else style = styleDateTime;
            return XlsxCell.Number(serial, style);
        }

        private XlsxCell TimeCell(TimeSpan v, ResultColumn col)
        {
            if (v < TimeSpan.Zero || v >= TimeSpan.FromDays(1)) return XlsxCell.Text(v.ToString("c", Inv));
            string serial = (v.Ticks / (double)TimeSpan.TicksPerDay).ToString("R", Inv);
            bool fraction = (v.Ticks % TimeSpan.TicksPerSecond) != 0;
            return XlsxCell.Number(serial, fraction ? styleTimeMs : styleTime);
        }

        private static string FormatDateTimeOffset(DateTimeOffset v, ResultColumn col)
        {
            int scale = col.Scale >= 0 && col.Scale <= 7 ? col.Scale : 7;
            string fmt = "yyyy-MM-dd HH:mm:ss" + (scale > 0 ? "." + new string('f', scale) : "") + " zzz";
            return v.ToString(fmt, Inv);
        }

        /// <summary>
        /// Text-Werte. Steht hinter der Spalte laut Schema ein Zahl- oder Datumstyp (z.B. weil
        /// SSMS den Wert als String gespeichert hat), wird kulturunabhängig zurückgewandelt.
        /// Echte Text-Spalten bleiben immer Text (führende Nullen, Artikelnummern, PLZ …).
        /// </summary>
        private XlsxCell StringValue(string s, ResultColumn col)
        {
            switch (col.SqlType)
            {
                case "tinyint":
                case "smallint":
                case "int":
                case "bigint":
                    if (long.TryParse(s, NumberStyles.AllowLeadingSign, Inv, out long l)) return Integer(l);
                    break;
                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                    if (SqlDecimalTryParse(s, out SqlDecimal d)) return Decimal(d, col);
                    break;
                case "float":
                case "real":
                    if (double.TryParse(s, NumberStyles.Float, Inv, out double f)) return Double(f);
                    break;
                case "date":
                case "datetime":
                case "datetime2":
                case "smalldatetime":
                    if (DateTime.TryParse(s, Inv, DateTimeStyles.None, out DateTime dt)) return DateTimeCell(dt, col);
                    break;
                case "time":
                    if (TimeSpan.TryParse(s, Inv, out TimeSpan ts)) return TimeCell(ts, col);
                    break;
            }
            return Text(s);
        }

        private static bool SqlDecimalTryParse(string s, out SqlDecimal value)
        {
            value = SqlDecimal.Null;
            if (string.IsNullOrWhiteSpace(s)) return false;
            try
            {
                value = SqlDecimal.Parse(s.Trim()); // erwartet invariantes Format ("." als Dezimaltrenner)
                return true;
            }
            catch (FormatException) { return false; }
            catch (OverflowException) { return false; }
        }

        private XlsxCell Text(string s)
        {
            s = XlsxWriter.SanitizeText(s ?? "");
            if (s.Length > XlsxWriter.MaxCellChars)
            {
                stats.TruncatedTexts++;
                s = s.Substring(0, XlsxWriter.MaxCellChars);
                if (char.IsHighSurrogate(s[s.Length - 1])) s = s.Substring(0, s.Length - 1);
            }
            return XlsxCell.Text(s);
        }

        private static string HexString(byte[] bytes)
        {
            if (bytes == null) return "";
            // Mehr als in eine Excel-Zelle passt, ist sinnlos; Text() kürzt (und zählt) den Rest.
            int len = Math.Min(bytes.Length, (XlsxWriter.MaxCellChars / 2) + 1);
            var sb = new StringBuilder(2 + len * 2);
            sb.Append("0x");
            for (int i = 0; i < len; i++) sb.Append(bytes[i].ToString("X2", Inv));
            return sb.ToString();
        }

        /// <summary>
        /// Anzahl der Stellen, die Excel exakt halten müsste: von der ersten Ziffer ≠ 0 bis
        /// zur letzten Ziffer; Nullen am Ende der Nachkommastellen zählen nicht, Nullen am Ende
        /// des Ganzzahlteils schon ("-0012.3400" → 4, "12300" → 5, "0.000123" → 3).
        /// Gleiche Regel wie die bigint-Grenze (> 15 Stellen → Text).
        /// </summary>
        internal static int SignificantDigits(string s)
        {
            if (string.IsNullOrEmpty(s)) return 1;
            int dot = s.IndexOf('.');
            if (dot >= 0) s = s.TrimEnd('0').TrimEnd('.');
            var digits = new StringBuilder(s.Length);
            foreach (char ch in s)
                if (ch >= '0' && ch <= '9') digits.Append(ch);
            string d = digits.ToString().TrimStart('0');
            return Math.Max(d.Length, 1);
        }

        /// <summary>
        /// Übersetzt das kurze Datumsformat der aktuellen Kultur in einen Excel-Formatcode
        /// (de-DE "dd.MM.yyyy" → "dd.mm.yyyy", en-US "M/d/yyyy" → "m/d/yyyy").
        /// Fallback ISO "yyyy-mm-dd".
        /// </summary>
        internal static string ExcelDatePattern(CultureInfo culture)
        {
            string p = culture?.DateTimeFormat?.ShortDatePattern;
            if (string.IsNullOrEmpty(p)) return "yyyy-mm-dd";
            var sb = new StringBuilder();
            bool hasD = false, hasM = false, hasY = false;
            for (int i = 0; i < p.Length; i++)
            {
                char ch = p[i];
                switch (ch)
                {
                    case 'd': sb.Append('d'); hasD = true; break;
                    case 'M': sb.Append('m'); hasM = true; break;
                    case 'y': sb.Append('y'); hasY = true; break;
                    case '.':
                    case '-':
                    case '/':
                    case ' ':
                        sb.Append(ch); break;
                    case '\'':
                        // Literal in .NET-Pattern → Excel-Literal in Anführungszeichen
                        int end = p.IndexOf('\'', i + 1);
                        if (end < 0) end = p.Length;
                        sb.Append('"').Append(p.Substring(i + 1, end - i - 1)).Append('"');
                        i = end;
                        break;
                    default:
                        sb.Append('\\').Append(ch); break;
                }
            }
            if (!(hasD && hasM && hasY)) return "yyyy-mm-dd";
            // Excel kennt "dddd" als Wochentag – kurze Datumsformate haben das nie, zur Sicherheit begrenzen.
            string result = sb.ToString().Replace("dddd", "dd").Replace("ddd", "dd");
            return result;
        }

        private static string UniqueSheetName(string wanted, HashSet<string> used)
        {
            var sb = new StringBuilder();
            foreach (char ch in wanted)
                sb.Append("[]:*?/\\".IndexOf(ch) >= 0 ? '_' : ch);
            string name = sb.ToString().Trim('\'');
            if (name.Length == 0) name = "Sheet";
            if (name.Length > 31) name = name.Substring(0, 31);
            string candidate = name;
            for (int n = 2; used.Contains(candidate); n++)
            {
                string suffix = " " + n.ToString(Inv);
                candidate = (name.Length + suffix.Length > 31 ? name.Substring(0, 31 - suffix.Length) : name) + suffix;
            }
            used.Add(candidate);
            return candidate;
        }
    }
}
