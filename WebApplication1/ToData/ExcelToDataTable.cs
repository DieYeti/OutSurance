using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WebApplication1.ToData
{
    public class ExcelToDataTable : IDisposable
    {
        #region Properties and Declares

        public SpreadsheetDocument CurrentDocument { get; private set; }
        public string CurrentFilename { get; private set; }
        public WorksheetPart CurrentSheet { get; set; }

        private static readonly List<string> ColumnNameList = new List<string>();
        private static readonly string[] ColumnNames = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                                                         "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
                                                         "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ",
                                                         "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ",
                                                         "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ",
                                                         "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ" };

        #endregion

        #region Construction

        public ExcelToDataTable(string fileName)
        {
            OpenDocument(fileName);
            OpenSheet();
        }

        #endregion

        #region Opening

        private void OpenDocument(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return;
            }

            SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileName, false);
            CurrentDocument = spreadSheet;
            CurrentFilename = fileName;
        }

        private void OpenSheet()
        {
            if (CurrentDocument == null)
            {
                return;
            }

            string ID = CurrentDocument.WorkbookPart.Workbook.Descendants<Sheet>().First().Id;
            CurrentSheet = (WorksheetPart)CurrentDocument.WorkbookPart.GetPartById(ID);
        }

        #endregion

        #region Reading

        public void ReadDocument(int columnCount, ref List<string[]> result)
        {
            var dataResult = new DataResult(DataResult.DataResultType.ListType);

            if (!ExecuteReadDocument(columnCount, ref dataResult))
            {
                return;
            }

            result = dataResult.GetList();
        }

        public void ReadDocument(int columnCount, ref DataTable result)
        {
            var dataResult = new DataResult(DataResult.DataResultType.DataTableType);

            if (!ExecuteReadDocument(columnCount, ref dataResult))
            {
                return;
            }

            result = dataResult.GetDataTable();
        }

        private bool ExecuteReadDocument(int columnCount, ref DataResult result)
        {
            if (CurrentSheet == null)
            {
                throw new Exception("No sheet selected");
            }

            var stringTableList = GetSharedStringPart().SharedStringTable.ChildElements.ToList();
            var lastRow = CurrentSheet.Worksheet.Descendants<Row>().LastOrDefault();

            if (lastRow == null)
            {
                return false;
            }

            Int16[] DateColumns = new Int16[columnCount];
            int cellCount = 0;

            var allRows = CurrentSheet.Worksheet.Descendants<Row>().ToList();

            for (var rowIndex = 1; rowIndex <= lastRow.RowIndex; rowIndex++)
            {
                if ((rowIndex != 2) && (rowIndex != 3))
                {
                    var cellList = new List<string>();
                    var cellValues = (from c in
                                          (from rows in allRows
                                           where rows.RowIndex.Value == rowIndex
                                           select rows).FirstOrDefault().Descendants<Cell>()
                                      where c.CellValue != null
                                      select c).ToList();

                    if (rowIndex == 1)
                    {
                        cellCount = cellValues.Count;
                    }

                    for (var cellIndex = 0; cellIndex < cellCount; cellIndex++)
                    {
                        var colName = GetColumnName(cellIndex);
                        var value = "";
                        var value2 = "";
                        var cell = (from c in cellValues
                                    where c.CellReference.Value.Equals(colName + rowIndex, StringComparison.CurrentCultureIgnoreCase)
                                    select c).FirstOrDefault();

                        if (cell != null)
                        {
                            int sharedStrIndex;
                            value = cell.InnerText;

                            if (int.TryParse(value, out sharedStrIndex))
                            {
                                if (sharedStrIndex < stringTableList.Count)
                                {
                                    value2 = stringTableList[sharedStrIndex].InnerText;
                                }
                            }
                        }

                        if (rowIndex == 1)
                        {
                            if (value2.Contains("Date"))
                            {
                                DateColumns[cellIndex] = 1;
                            }
                            else
                            {
                                DateColumns[cellIndex] = 0;
                            }
                        }

                        try
                        {
                            if (cell.DataType == "s")
                            {
                                value = value2;
                            }
                            else if ((rowIndex != 1) && (DateColumns[cellIndex] == 1))
                            {
                                value = DateTime.FromOADate(Int32.Parse(value)).ToString();
                            }
                        }
                        catch { }

                        cellList.Add(value);
                    }

                    if (rowIndex == 1)
                    {
                        cellList.Add("RowID");
                    }
                    else
                    {
                        cellList.Add(rowIndex.ToString());
                    }

                    result.AddRow(cellList.ToArray());
                }
            }

            return true;
        }

        private SharedStringTablePart GetSharedStringPart()
        {
            return CurrentDocument.WorkbookPart.SharedStringTablePart;
        }

        public static string GetColumnName(int colIndex)
        {
            if (colIndex < 0)
            {
                return "#";
            }

            if (ColumnNameList.Count <= colIndex)
            {
                for (int index = ColumnNameList.Count; index < (colIndex + 1); index++)
                {
                    string colName;

                    if (index >= ColumnNames.Length)
                    {
                        var subIndex = (int)Math.Floor((double)index / ColumnNames.Length) - 1;
                        int sufIndex = (index - ((subIndex + 1) * ColumnNames.Length));
                        colName = GetColumnName(subIndex) + GetColumnName(sufIndex);
                    }
                    else
                    {
                        colName = ColumnNames[index];
                    }

                    ColumnNameList.Add(colName);
                }
            }

            return ColumnNameList[colIndex];
        }

        #endregion

        #region Disposing

        public void Dispose()
        {
            if (CurrentDocument != null)
            {
                CurrentDocument.Dispose();
            }

            CurrentSheet = null;
            CurrentDocument = null;
            CurrentFilename = null;
        }

        #endregion
    }

    #region Shared class

    public class DataResult
    {
        public enum DataResultType { DataTableType, ListType }
        private readonly DataResultType _dataResultType;
        private readonly DataTable _table;
        private readonly List<string[]> _list;

        public DataResult(DataResultType resultType)
        {
            _dataResultType = resultType;

            switch (resultType)
            {
                case DataResultType.DataTableType:
                    _table = new DataTable();
                    break;
                case DataResultType.ListType:
                    _list = new List<string[]>();
                    break;
                default:
                    throw new ArgumentOutOfRangeException("DataResult.DataResultType does not exist");
            }
        }

        public void AddRow(string[] rowData)
        {
            switch (_dataResultType)
            {
                case DataResultType.DataTableType:
                    while (_table.Columns.Count < rowData.Length)
                    {
                        _table.Columns.Add(ExcelToDataTable.GetColumnName(_table.Columns.Count));
                    }

                    _table.Rows.Add(rowData);
                    break;
                case DataResultType.ListType:
                    _list.Add(rowData);
                    break;
                default:
                    throw new ArgumentOutOfRangeException("DataResult.DataResultType does not exist");
            }
        }

        public DataTable GetDataTable()
        {
            return _table;
        }

        public List<string[]> GetList()
        {
            return _list;
        }
    }

    #endregion
}