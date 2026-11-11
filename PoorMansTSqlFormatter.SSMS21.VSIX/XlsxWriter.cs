using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml;

namespace PoorMansTSqlFormatter.SSMS21.VSIX
{
    /// <summary>Art einer Excel-Zelle.</summary>
    internal enum XlsxCellKind
    {
        Empty,
        Number,
        Text,
    }

    /// <summary>
    /// Eine fertig konvertierte Zelle. <see cref="Value"/> ist bei Zahlen bereits invariant
    /// formatiert (Punkt als Dezimaltrenner) – so verlangt es das Dateiformat, unabhängig von
    /// der Ländereinstellung. Excel zeigt die Zahl später in der Kultur des Anwenders an.
    /// </summary>
    internal struct XlsxCell
    {
        public XlsxCellKind Kind;
        public string Value;
        public int Style;

        public static readonly XlsxCell Empty = new XlsxCell { Kind = XlsxCellKind.Empty };

        public static XlsxCell Number(string invariantNumber, int style) =>
            new XlsxCell { Kind = XlsxCellKind.Number, Value = invariantNumber, Style = style };

        public static XlsxCell Text(string text, int style = 0) =>
            new XlsxCell { Kind = XlsxCellKind.Text, Value = text, Style = style };
    }

    /// <summary>
    /// Minimaler, streamender Writer für .xlsx (Office Open XML SpreadsheetML), ohne
    /// externe Abhängigkeiten (kein Excel-Interop, kein OpenXML-SDK). Arbeitsblätter werden
    /// nacheinander geschrieben; Workbook, Styles und Content-Types erst beim Dispose,
    /// weil erst dann alle Blätter und Zahlenformate bekannt sind.
    /// </summary>
    internal sealed class XlsxWriter : IDisposable
    {
        /// <summary>Excel: max. 1.048.576 Zeilen je Blatt (inkl. Kopfzeile).</summary>
        public const int MaxRowsPerSheet = 1048576;

        /// <summary>Excel: max. 32.767 Zeichen je Zelle.</summary>
        public const int MaxCellChars = 32767;

        /// <summary>Stil-Index der fetten Kopfzeile.</summary>
        public const int HeaderStyle = 1;

        private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private readonly FileStream file;
        private readonly ZipArchive zip;
        private readonly List<SheetInfo> sheets = new List<SheetInfo>();
        private readonly List<string> customFormats = new List<string>();
        private readonly Dictionary<string, int> styleByFormat = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<int> styleNumFmtIds = new List<int>();
        private bool disposed;

        private sealed class SheetInfo
        {
            public string Name;
            public string FilterRange;
        }

        public XlsxWriter(string path)
        {
            file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            zip = new ZipArchive(file, ZipArchiveMode.Create, false, Encoding.UTF8);
        }

        /// <summary>
        /// Liefert den Stil-Index für ein Excel-Zahlenformat (z.B. "0", "#,##0.00",
        /// "dd.mm.yyyy hh:mm:ss"). Formatcodes im Dateiformat sind immer englisch notiert
        /// (Punkt = Dezimal, Komma = Tausender); Excel lokalisiert die Anzeige selbst.
        /// </summary>
        public int GetStyle(string formatCode)
        {
            if (string.IsNullOrEmpty(formatCode) || formatCode == "General") return 0;
            if (styleByFormat.TryGetValue(formatCode, out int style)) return style;

            int numFmtId;
            if (formatCode == "0") numFmtId = 1;           // eingebaut
            else if (formatCode == "0.00") numFmtId = 2;   // eingebaut
            else
            {
                customFormats.Add(formatCode);
                numFmtId = 163 + customFormats.Count;      // eigene Formate ab 164
            }
            styleNumFmtIds.Add(numFmtId);
            style = 1 + styleNumFmtIds.Count;              // 0 = Standard, 1 = Kopfzeile
            styleByFormat[formatCode] = style;
            return style;
        }

        /// <summary>Beginnt ein neues Arbeitsblatt. Es muss vor dem nächsten geschlossen werden.</summary>
        public SheetWriter BeginSheet(string name, double[] columnWidths)
        {
            var info = new SheetInfo { Name = name };
            sheets.Add(info);
            ZipArchiveEntry entry = zip.CreateEntry("xl/worksheets/sheet" + sheets.Count + ".xml", CompressionLevel.Fastest);
            return new SheetWriter(entry.Open(), columnWidths, sheets.Count == 1, range => info.FilterRange = range);
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            try
            {
                WriteEntry("[Content_Types].xml", BuildContentTypes());
                WriteEntry("_rels/.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"" + NsRel + "/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                WriteEntry("xl/workbook.xml", BuildWorkbook());
                WriteEntry("xl/_rels/workbook.xml.rels", BuildWorkbookRels());
                WriteEntry("xl/styles.xml", BuildStyles());
            }
            finally
            {
                zip.Dispose();
                file.Dispose();
            }
        }

        private void WriteEntry(string name, string content)
        {
            ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.Fastest);
            using (var w = new StreamWriter(entry.Open(), new UTF8Encoding(false)))
            {
                w.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
                w.Write(content);
            }
        }

        private string BuildContentTypes()
        {
            var sb = new StringBuilder();
            sb.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            sb.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            for (int i = 1; i <= sheets.Count; i++)
                sb.Append("<Override PartName=\"/xl/worksheets/sheet").Append(i)
                  .Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        private string BuildWorkbook()
        {
            var sb = new StringBuilder();
            sb.Append("<workbook xmlns=\"").Append(NsMain).Append("\" xmlns:r=\"").Append(NsRel).Append("\">");
            sb.Append("<bookViews><workbookView/></bookViews><sheets>");
            for (int i = 0; i < sheets.Count; i++)
                sb.Append("<sheet name=\"").Append(Esc(sheets[i].Name)).Append("\" sheetId=\"").Append(i + 1)
                  .Append("\" r:id=\"rId").Append(i + 1).Append("\"/>");
            sb.Append("</sheets>");

            // Der AutoFilter braucht den versteckten Namen _FilterDatabase je Blatt.
            bool any = false;
            for (int i = 0; i < sheets.Count; i++)
            {
                if (sheets[i].FilterRange == null) continue;
                if (!any) { sb.Append("<definedNames>"); any = true; }
                sb.Append("<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"").Append(i)
                  .Append("\" hidden=\"1\">'").Append(Esc(sheets[i].Name.Replace("'", "''"))).Append("'!")
                  .Append(Esc(sheets[i].FilterRange)).Append("</definedName>");
            }
            if (any) sb.Append("</definedNames>");
            sb.Append("</workbook>");
            return sb.ToString();
        }

        private string BuildWorkbookRels()
        {
            var sb = new StringBuilder();
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 1; i <= sheets.Count; i++)
                sb.Append("<Relationship Id=\"rId").Append(i).Append("\" Type=\"").Append(NsRel)
                  .Append("/worksheet\" Target=\"worksheets/sheet").Append(i).Append(".xml\"/>");
            sb.Append("<Relationship Id=\"rId").Append(sheets.Count + 1).Append("\" Type=\"").Append(NsRel)
              .Append("/styles\" Target=\"styles.xml\"/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private string BuildStyles()
        {
            var sb = new StringBuilder();
            sb.Append("<styleSheet xmlns=\"").Append(NsMain).Append("\">");
            if (customFormats.Count > 0)
            {
                sb.Append("<numFmts count=\"").Append(customFormats.Count).Append("\">");
                for (int i = 0; i < customFormats.Count; i++)
                    sb.Append("<numFmt numFmtId=\"").Append(164 + i).Append("\" formatCode=\"").Append(Esc(customFormats[i])).Append("\"/>");
                sb.Append("</numFmts>");
            }
            sb.Append("<fonts count=\"2\">");
            sb.Append("<font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font>");
            sb.Append("<font><b/><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font>");
            sb.Append("</fonts>");
            sb.Append("<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>");
            sb.Append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFD9E1F2\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>");
            sb.Append("<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>");
            sb.Append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
            sb.Append("<cellXfs count=\"").Append(2 + styleNumFmtIds.Count).Append("\">");
            sb.Append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>");
            sb.Append("<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/>");
            foreach (int id in styleNumFmtIds)
                sb.Append("<xf numFmtId=\"").Append(id).Append("\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>");
            sb.Append("</cellXfs>");
            sb.Append("<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>");
            sb.Append("</styleSheet>");
            return sb.ToString();
        }

        private static string Esc(string s) => SecurityElement.Escape(s) ?? "";

        /// <summary>Spaltenbuchstaben (0 → A, 25 → Z, 26 → AA …).</summary>
        public static string ColumnName(int index)
        {
            var sb = new StringBuilder();
            index++;
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                index = (index - 1) / 26;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Entfernt Zeichen, die in XML 1.0 nicht erlaubt sind (Steuerzeichen außer Tab/LF/CR,
        /// einzelne Surrogate) und vereinheitlicht CRLF zu LF (Zeilenumbruch in der Zelle).
        /// </summary>
        public static string SanitizeText(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            StringBuilder sb = null;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                bool ok;
                if (c == '\r' && i + 1 < s.Length && s[i + 1] == '\n') ok = false;   // CR von CRLF weglassen
                else if (c == '\t' || c == '\n' || c == '\r') ok = true;
                else if (c < 0x20 || c == 0xFFFE || c == 0xFFFF) ok = false;
                else if (char.IsHighSurrogate(c)) ok = i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]);
                else if (char.IsLowSurrogate(c)) ok = i > 0 && char.IsHighSurrogate(s[i - 1]);
                else ok = true;

                if (ok) sb?.Append(c);
                else if (sb == null) sb = new StringBuilder(s, 0, i, s.Length);
            }
            return sb?.ToString() ?? s;
        }

        /// <summary>Schreibt die Zeilen eines Arbeitsblatts (streamend).</summary>
        internal sealed class SheetWriter : IDisposable
        {
            private readonly XmlWriter xml;
            private readonly Action<string> setFilterRange;
            private readonly int columnCount;
            private readonly string[] columnNames;
            private int rowNumber;

            internal SheetWriter(Stream stream, double[] columnWidths, bool selected, Action<string> setFilterRange)
            {
                this.setFilterRange = setFilterRange;
                columnCount = columnWidths.Length;
                columnNames = new string[columnCount];
                for (int i = 0; i < columnCount; i++) columnNames[i] = ColumnName(i);

                xml = XmlWriter.Create(stream, new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    CheckCharacters = false, // Texte werden vorab bereinigt (SanitizeText)
                    CloseOutput = true,
                });
                xml.WriteStartDocument(true);
                xml.WriteStartElement("worksheet", NsMain);

                // Kopfzeile fixieren
                xml.WriteStartElement("sheetViews");
                xml.WriteStartElement("sheetView");
                if (selected) xml.WriteAttributeString("tabSelected", "1");
                xml.WriteAttributeString("workbookViewId", "0");
                xml.WriteStartElement("pane");
                xml.WriteAttributeString("ySplit", "1");
                xml.WriteAttributeString("topLeftCell", "A2");
                xml.WriteAttributeString("activePane", "bottomLeft");
                xml.WriteAttributeString("state", "frozen");
                xml.WriteEndElement();
                xml.WriteStartElement("selection");
                xml.WriteAttributeString("pane", "bottomLeft");
                xml.WriteAttributeString("activeCell", "A2");
                xml.WriteAttributeString("sqref", "A2");
                xml.WriteEndElement();
                xml.WriteEndElement();
                xml.WriteEndElement();

                if (columnCount > 0)
                {
                    xml.WriteStartElement("cols");
                    for (int i = 0; i < columnCount; i++)
                    {
                        xml.WriteStartElement("col");
                        xml.WriteAttributeString("min", (i + 1).ToString(CultureInfo.InvariantCulture));
                        xml.WriteAttributeString("max", (i + 1).ToString(CultureInfo.InvariantCulture));
                        xml.WriteAttributeString("width", columnWidths[i].ToString("0.##", CultureInfo.InvariantCulture));
                        xml.WriteAttributeString("customWidth", "1");
                        xml.WriteEndElement();
                    }
                    xml.WriteEndElement();
                }
                xml.WriteStartElement("sheetData");
            }

            public void WriteRow(XlsxCell[] cells)
            {
                rowNumber++;
                string r = rowNumber.ToString(CultureInfo.InvariantCulture);
                xml.WriteStartElement("row");
                xml.WriteAttributeString("r", r);
                for (int i = 0; i < cells.Length && i < columnCount; i++)
                {
                    XlsxCell cell = cells[i];
                    if (cell.Kind == XlsxCellKind.Empty) continue;

                    xml.WriteStartElement("c");
                    xml.WriteAttributeString("r", columnNames[i] + r);
                    if (cell.Style != 0) xml.WriteAttributeString("s", cell.Style.ToString(CultureInfo.InvariantCulture));
                    if (cell.Kind == XlsxCellKind.Number)
                    {
                        xml.WriteElementString("v", cell.Value);
                    }
                    else
                    {
                        // Inline-String: wird NIE als Zahl, Datum oder Formel interpretiert
                        // (führende Nullen, "=…", "1-2" usw. bleiben exakt erhalten).
                        xml.WriteAttributeString("t", "inlineStr");
                        xml.WriteStartElement("is");
                        xml.WriteStartElement("t");
                        string v = cell.Value ?? "";
                        if (v.Length > 0 && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[v.Length - 1])))
                            xml.WriteAttributeString("xml", "space", null, "preserve");
                        xml.WriteString(v);
                        xml.WriteEndElement();
                        xml.WriteEndElement();
                    }
                    xml.WriteEndElement();
                }
                xml.WriteEndElement();
            }

            public void Dispose()
            {
                xml.WriteEndElement(); // sheetData
                if (columnCount > 0)
                {
                    string range = "$A$1:$" + columnNames[columnCount - 1] + "$" + Math.Max(rowNumber, 1).ToString(CultureInfo.InvariantCulture);
                    xml.WriteStartElement("autoFilter");
                    xml.WriteAttributeString("ref", range.Replace("$", ""));
                    xml.WriteEndElement();
                    setFilterRange(range);
                }
                xml.WriteStartElement("pageMargins");
                xml.WriteAttributeString("left", "0.7");
                xml.WriteAttributeString("right", "0.7");
                xml.WriteAttributeString("top", "0.78740157499999996");
                xml.WriteAttributeString("bottom", "0.78740157499999996");
                xml.WriteAttributeString("header", "0.3");
                xml.WriteAttributeString("footer", "0.3");
                xml.WriteEndElement();
                xml.WriteEndElement(); // worksheet
                xml.WriteEndDocument();
                xml.Dispose();
            }
        }
    }
}
