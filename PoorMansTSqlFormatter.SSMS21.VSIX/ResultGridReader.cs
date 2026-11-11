using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace PoorMansTSqlFormatter.SSMS21.VSIX
{
    /// <summary>Metadaten einer Ergebnisspalte (aus dem Schema des SqlDataReaders).</summary>
    internal sealed class ResultColumn
    {
        public string Name;

        /// <summary>SQL-Server-Typname in Kleinbuchstaben (z.B. "int", "decimal", "datetime2", "nvarchar").</summary>
        public string SqlType;

        /// <summary>NumericScale aus dem Schema (Nachkommastellen bzw. Sekundenbruchteile), -1 = unbekannt.</summary>
        public int Scale = -1;
    }

    /// <summary>
    /// Ein Ergebnis-Raster aus dem SSMS-Ergebnisbereich. Kapselt das interne SSMS-Objekt
    /// QEResultSet (SqlEditors.dll), auf das nur per Reflection zugegriffen werden kann.
    /// </summary>
    internal sealed class ResultSetSnapshot
    {
        private readonly object resultSet;
        private readonly MethodInfo getCellData;
        private readonly int columnOffset;

        public ResultColumn[] Columns { get; }

        /// <summary>Anzahl der Zeilen zum Zeitpunkt des Exports.</summary>
        public long RowCount { get; }

        /// <summary>False, wenn die Abfrage noch lief und nicht alle Zeilen geladen waren.</summary>
        public bool Complete { get; }

        internal ResultSetSnapshot(object resultSet, MethodInfo getCellData, int columnOffset,
            ResultColumn[] columns, long rowCount, bool complete)
        {
            this.resultSet = resultSet;
            this.getCellData = getCellData;
            this.columnOffset = columnOffset;
            Columns = columns;
            RowCount = rowCount;
            Complete = complete;
        }

        /// <summary>
        /// Liefert den typisierten Zellwert, so wie ihn SSMS gespeichert hat (z.B. SqlInt32,
        /// SqlDecimal, SqlDateTime, DateTime, TimeSpan, DateTimeOffset, SqlGuid, SqlString,
        /// byte[] …) – NICHT den angezeigten Text. Spaltenindex 0-basiert über die Datenspalten.
        /// </summary>
        public object GetValue(long row, int column)
        {
            return getCellData.Invoke(resultSet, new object[] { row, column + columnOffset });
        }
    }

    /// <summary>
    /// Liest die Ergebnisraster des aktiven SSMS-Abfragefensters aus.
    /// Kette (SSMS 22, SqlEditors.dll, per IL verifiziert):
    /// SqlScriptEditorControl.m_sqlResultsControl (DisplaySQLResultsControl)
    ///   .m_gridResultsPage (GridResultsTabPage)
    ///   .NumberOfGrids / GetResultSet(i) (QEResultSet)
    ///   .GetCellData(row, col) / GetSchemaRow(col) / TotalNumberOfRows / GridMode.
    /// Im Grid-Modus ist Spalte 0 die Zeilennummern-Spalte, deshalb Offset +1 bei GetCellData;
    /// GetSchemaRow ist 0-basiert über die Datenspalten.
    /// </summary>
    internal static class ResultGridReader
    {
        private const BindingFlags AnyInstance =
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

        private const BindingFlags AnyStatic =
            BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

        /// <summary>
        /// Ermittelt alle Ergebnisraster des aktiven Abfragefensters. Leere Liste, wenn kein
        /// Abfragefenster aktiv ist oder es keine Raster-Ergebnisse gibt (z.B. „Ergebnisse in Text").
        /// </summary>
        public static List<ResultSetSnapshot> ReadActiveWindow()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var result = new List<ResultSetSnapshot>();

            object resultsControl = FindResultsControl();
            object gridPage = GetMember(resultsControl, "m_gridResultsPage");
            if (gridPage == null) return result;

            foreach (object rs in EnumerateResultSets(gridPage))
            {
                ResultSetSnapshot snapshot = CreateSnapshot(rs);
                if (snapshot != null) result.Add(snapshot);
            }
            return result;
        }

        /// <summary>
        /// Sucht das DisplaySQLResultsControl des aktiven Abfragefensters. Zuerst über den
        /// aktiven Dokument-Frame (funktioniert auch, wenn der Fokus im Ergebnisraster oder im
        /// Objekt-Explorer steht), dann über die SSMS-eigenen Hilfsmethoden als Fallback.
        /// </summary>
        private static object FindResultsControl()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var monSel = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            if (monSel != null &&
                monSel.GetCurrentElementValue((uint)VSConstants.VSSELELEMID.SEID_DocumentFrame, out object frameObj) == VSConstants.S_OK &&
                frameObj is IVsWindowFrame frame &&
                frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) == VSConstants.S_OK)
            {
                object control = GetMember(docView, "m_sqlResultsControl");
                if (control != null) return control;
            }

            // Fallback 1: ScriptFactory.Instance.GetCurrentlyActiveFrameDocView(monSel, false, out frame)
            Type scriptFactoryType = FindSqlEditorsType("Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptFactory");
            if (scriptFactoryType != null && monSel != null)
            {
                try
                {
                    object factory = scriptFactoryType.GetProperty("Instance", AnyStatic)?.GetValue(null);
                    MethodInfo m = scriptFactoryType.GetMethod("GetCurrentlyActiveFrameDocView", AnyInstance);
                    if (factory != null && m != null)
                    {
                        object activeView = m.Invoke(factory, new object[] { monSel, false, null });
                        object control = GetMember(activeView, "m_sqlResultsControl");
                        if (control != null) return control;
                    }
                }
                catch (Exception ex)
                {
                    Log("ScriptFactory-Fallback: " + ex.Message);
                }
            }

            // Fallback 2: ScriptEditorControl.GetActiveScriptEditorControl()
            Type editorType = FindSqlEditorsType("Microsoft.SqlServer.Management.UI.VSIntegration.Editors.ScriptEditorControl");
            if (editorType != null)
            {
                try
                {
                    object editor = editorType.GetMethod("GetActiveScriptEditorControl", AnyStatic)?.Invoke(null, null);
                    return GetMember(editor, "m_sqlResultsControl");
                }
                catch (Exception ex)
                {
                    Log("GetActiveScriptEditorControl-Fallback: " + ex.Message);
                }
            }
            return null;
        }

        /// <summary>Alle QEResultSet-Objekte der Raster-Seite, in Anzeigereihenfolge.</summary>
        private static IEnumerable<object> EnumerateResultSets(object gridPage)
        {
            Type t = gridPage.GetType();
            PropertyInfo countProp = t.GetProperty("NumberOfGrids", AnyInstance);
            MethodInfo getResultSet = t.GetMethod("GetResultSet", AnyInstance, null, new[] { typeof(int) }, null);

            if (countProp != null && getResultSet != null)
            {
                int count = (int)countProp.GetValue(gridPage);
                for (int i = 0; i < count; i++)
                {
                    object rs = getResultSet.Invoke(gridPage, new object[] { i });
                    if (rs != null) yield return rs;
                }
                yield break;
            }

            // Fallback: Liste der ResultSetAndGridContainer direkt lesen.
            if (GetMember(gridPage, "m_gridContainers") is IEnumerable containers)
            {
                foreach (object container in containers)
                {
                    object rs = GetMember(container, "QEResultSet") ?? GetMember(container, "m_rs");
                    if (rs != null) yield return rs;
                }
            }
        }

        private static ResultSetSnapshot CreateSnapshot(object rs)
        {
            Type t = rs.GetType();
            MethodInfo getCellData = t.GetMethod("GetCellData", AnyInstance, null, new[] { typeof(long), typeof(int) }, null);
            MethodInfo getSchemaRow = t.GetMethod("GetSchemaRow", AnyInstance, null, new[] { typeof(int) }, null);
            if (getCellData == null || getSchemaRow == null)
            {
                Log("QEResultSet ohne GetCellData/GetSchemaRow – SSMS-Version inkompatibel?");
                return null;
            }

            int columnCount = Convert.ToInt32(GetMember(rs, "NumberOfDataColumns") ?? 0);
            long rowCount = Convert.ToInt64(GetMember(rs, "TotalNumberOfRows") ?? 0L);
            bool gridMode = GetMember(rs, "GridMode") as bool? ?? true;
            bool complete = GetMember(rs, "StoredAllData") as bool? ?? true;

            var names = GetMember(rs, "ColumnNames") as IList;
            var columns = new ResultColumn[columnCount];
            for (int c = 0; c < columnCount; c++)
            {
                var col = new ResultColumn
                {
                    Name = names != null && c < names.Count ? names[c] as string : null,
                    SqlType = ""
                };
                try
                {
                    if (getSchemaRow.Invoke(rs, new object[] { c }) is DataRow schema)
                    {
                        if (string.IsNullOrEmpty(col.Name)) col.Name = SchemaValue(schema, "ColumnName") as string;
                        col.SqlType = (SchemaValue(schema, "DataTypeName") as string ?? "").ToLowerInvariant();
                        object scale = SchemaValue(schema, "NumericScale");
                        if (scale != null && !(scale is DBNull)) col.Scale = Convert.ToInt32(scale);
                    }
                }
                catch (Exception ex)
                {
                    Log("Schema Spalte " + c + ": " + ex.Message);
                }
                if (string.IsNullOrEmpty(col.Name)) col.Name = Localizer.T("(Kein Spaltenname)", "(No column name)");
                columns[c] = col;
            }

            return new ResultSetSnapshot(rs, getCellData, gridMode ? 1 : 0, columns, rowCount, complete);
        }

        private static object SchemaValue(DataRow row, string column)
        {
            return row.Table.Columns.Contains(column) ? row[column] : null;
        }

        /// <summary>Liest Feld oder Eigenschaft (auch nicht-öffentlich, auch aus Basisklassen).</summary>
        private static object GetMember(object instance, string name)
        {
            if (instance == null) return null;
            for (Type t = instance.GetType(); t != null; t = t.BaseType)
            {
                FieldInfo f = t.GetField(name, AnyInstance | BindingFlags.DeclaredOnly);
                if (f != null) return f.GetValue(instance);
                PropertyInfo p = t.GetProperty(name, AnyInstance | BindingFlags.DeclaredOnly);
                if (p != null && p.GetIndexParameters().Length == 0) return p.GetValue(instance);
            }
            return null;
        }

        private static Type FindSqlEditorsType(string fullName)
        {
            foreach (Assembly a in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (!string.Equals(a.GetName().Name, "SqlEditors", StringComparison.OrdinalIgnoreCase)) continue;
                Type t = a.GetType(fullName, false);
                if (t != null) return t;
            }
            return null;
        }

        internal static void Log(string message) => FormatSqlCommand.Log("Excel-Export: " + message);
    }
}
