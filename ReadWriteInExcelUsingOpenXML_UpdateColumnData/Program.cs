using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ReadWriteInExcelUsingOpenXML_UpdateColumnData
{
    public partial class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string destinationPath = @"C:\Users\Vishal Yelve\VishalTest_Destination";
                var _dataTable = GetExcelDatatoDataTable();
                if (_dataTable.Rows.Count > 0)
                {
                    foreach (DataRow row in _dataTable.Rows)
                    {
                        string FilePath = row.ItemArray[1].ToString();
                        if (File.Exists(FilePath))
                        {
                            row.SetField("File New Name", DateTime.Now.ToString("yyyyMMddHHmmssff") + Path.GetExtension(FilePath));
                            Console.WriteLine("File New Name:" + DateTime.Now.ToString("yyyyMMddHHmmssff") + Path.GetExtension(FilePath));

                            //File.Copy(FilePath, @"C:\Users\Vishal Yelve\VishalTest_Destination\" + DateTime.Now.ToString("yyyyMMddHHmmssff") + Path.GetExtension(FilePath));

                            Copy(Path.GetDirectoryName(FilePath), destinationPath); 
                        }
                        else
                        {
                            row.SetField("File New Name", string.Empty);
                            Console.WriteLine("File New Name:" + string.Empty);
                        }
                    }

                    WriteExcelFile(_dataTable, @"C:\Users\Vishal Yelve\ReadDatafromExcel.xlsx");
                }

                Console.WriteLine("Excel Updated!!");

            }
            catch (Exception)
            {
                Console.WriteLine("Error occure while processing");
                Console.ReadKey();
            }

        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }

        private static void WriteExcelFile(DataTable table, string outputPath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }

        private static DataTable GetExcelDatatoDataTable()
        {
            var table = new DataTable();
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(@"C:\Users\Vishal Yelve\ReadDatafromExcel.xlsx", false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    table.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in rows) //this will also include your header row...
                {
                    DataRow tempRow = table.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    table.Rows.Add(tempRow);
                }
            }

            table.Rows.RemoveAt(0);

            return table;
        }


        public static void Copy(string sourceDirectory, string targetDirectory)
        {
            var diSource = new DirectoryInfo(sourceDirectory);
            var diTarget = new DirectoryInfo(targetDirectory);

            CopyAll(diSource, diTarget);
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                //fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);

                fi.CopyTo(Path.Combine(target.FullName, DateTime.Now.ToString("yyyyMMddHHmmssff") + Path.GetExtension(fi.Name)), true);

            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir = target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }
        }
    }
}
