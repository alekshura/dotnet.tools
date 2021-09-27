using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Dotnet.Tools.OpenXml.Excel
{
    /// <summary>
    /// Helper class for reading Excel in OpenXml format 
    /// </summary>
    public class ExcelHelper
    {
        /// <summary>
        /// Method reads Excel file nd returns the <see cref="DataSet"/> for data in excel. Every sheet corresponds to one table.
        /// </summary>
        /// <param name="fileName">The full path to the excel file</param>
        /// <returns></returns>
        public static DataSet CachExcelFile(string fileName)
        {
            try
            {
                var dataSet = new DataSet();
                using var doc = SpreadsheetDocument.Open(fileName, false);
                var workbookPart = doc.WorkbookPart;
                var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

                foreach (var sheet in sheets.OfType<Sheet>())
                {
                    var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    var dataTable = new DataTable(sheet.Name);

                    for (int rowCounter = 0; rowCounter < sheetData.ChildElements.Count; rowCounter++)
                    {
                        List<string> rowList = new();
                        for (int columnCounter = 0; columnCounter < sheetData.ElementAt(rowCounter).ChildElements.Count; columnCounter++)
                        {
                            var cell = (Cell)sheetData.ElementAt(rowCounter).ChildElements.ElementAt(columnCounter);

                            string cellValue = string.Empty;
                            if (cell.DataType != null)
                            {
                                if (cell.DataType == CellValues.SharedString)
                                {
                                    if (int.TryParse(cell.InnerText, out int id))
                                    {
                                        SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                        if (item.Text != null)
                                        {
                                            cellValue = item.Text.Text;
                                        }
                                        else if (item.InnerText != null)
                                        {
                                            cellValue = item.InnerText;
                                        }
                                        else if (item.InnerXml != null)
                                        {
                                            cellValue = item.InnerXml;
                                        }
                                        //Column name
                                        if (rowCounter == 0)
                                        {
                                            dataTable.Columns.Add(cellValue);
                                        }
                                        else
                                        {
                                            rowList.Add(cellValue);
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (rowCounter != 0)//reserved for column values
                                {
                                    rowList.Add(cell.InnerText);
                                }
                            }
                        }
                        if (rowCounter != 0)//reserved for column values
                            dataTable.Rows.Add(rowList.ToArray());

                    }

                    dataSet.Tables.Add(dataTable);
                }
                return dataSet;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }
    }
}
