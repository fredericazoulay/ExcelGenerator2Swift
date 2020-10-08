using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;

namespace ExcelGenerator2SwiftApp
{
    public static class HelperConvert
    {

            // remove "this" if not on C# 3.0 / .NET 3.5
            /*
            public static DataTable ToDataTable<T>(this IList<T> data)
            {
                PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
                DataTable table = new DataTable();
                for (int i = 0; i < props.Count; i++)
                {
                    PropertyDescriptor prop = props[i];
                    table.Columns.Add(prop.Name, prop.PropertyType);
                }
                object[] values = new object[props.Count];
                foreach (T item in data)
                {
                    for (int i = 0; i < values.Length; i++)
                    {
                        values[i] = props[i].GetValue(item);
                    }
                    table.Rows.Add(values);
                }
                return table;
            }
        */

            /*
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }
        */


            public static DataTable ToDataTable<T>(this IEnumerable<T> collection, string tableName)
            {
                DataTable tbl = ToDataTable(collection);
                tbl.TableName = tableName;
                return tbl;
            }

            public static DataTable ToDataTable<T>(this IEnumerable<T> collection)
            {
                DataTable dt = new DataTable();
                Type t = typeof(T);
                PropertyInfo[] pia = t.GetProperties();
                object temp;
                DataRow dr;

                for (int i = 0; i < pia.Length; i++)
                {
                    dt.Columns.Add(pia[i].Name, Nullable.GetUnderlyingType(pia[i].PropertyType) ?? pia[i].PropertyType);
                    dt.Columns[i].AllowDBNull = true;
                }

                //Populate the table
                foreach (T item in collection)
                {
                    dr = dt.NewRow();
                    dr.BeginEdit();

                    for (int i = 0; i < pia.Length; i++)
                    {
                        temp = pia[i].GetValue(item, null);
                        if (temp == null || (temp.GetType().Name == "Char" && ((char)temp).Equals('\0')))
                        {
                            dr[pia[i].Name] = (object)DBNull.Value;
                        }
                        else
                        {
                            dr[pia[i].Name] = temp;
                        }
                    }

                    dr.EndEdit();
                    dt.Rows.Add(dr);
                }
                return dt;
            }


        public static void ExportExcelFile(DataTable dataTable, string stFileName)
        {
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Sheets[1];

            //DataTable dataTable = new DataTable();
            //DataColumn column = new DataColumn("My Datacolumn");
            //
            //dataTable.Columns.Add(column);
            //dataTable.Rows.Add(new object[] { "Foobar" });

            var columns = dataTable.Columns.Count;
            var rows = dataTable.Rows.Count;

            Microsoft.Office.Interop.Excel.Range range = worksheet.Range["A1", String.Format("{0}{1}", GetExcelColumnName(columns), rows)];

            object[,] data = new object[rows, columns];

            for (int rowNumber = 0; rowNumber < rows; rowNumber++)
            {
                for (int columnNumber = 0; columnNumber < columns; columnNumber++)
                {
                    data[rowNumber, columnNumber] = dataTable.Rows[rowNumber][columnNumber].ToString();
                }
            }

            range.Value = data;

            workbook.SaveAs(@stFileName);
            workbook.Close();

            Marshal.ReleaseComObject(application);
            return;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        /*
        public void ExportDataSet(DataSet ds, string destination)
        {
            using (var workbook = SpreadsheetDocument.Create(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                foreach (System.Data.DataTable table in ds.Tables)
                {

                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    uint sheetId = 1;
                    if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<String> columns = new List<string>();
                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }


                    sheetData.AppendChild(headerRow);

                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                }
            }
        }
        */


        /*
 public DataTable ConvertListToDataTable<T>(IList<T> data)
 {
     PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
     DataTable table = new DataTable();
     foreach (PropertyDescriptor prop in properties)
         table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
     foreach (T item in data)
     {
         DataRow row = table.NewRow();
         foreach (PropertyDescriptor prop in properties)
             row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
         table.Rows.Add(row);
     }
     return table;

 }
 */

        /*
        static DataTable ConvertListToDataTable(List<string> list)
        {
            // New table.
            DataTable table = new DataTable();

            // Get max columns.
            int columns = 0;
            foreach (var array in list)
            {
                if (array.Length > columns)
                {
                    columns = array.Length;
                }
            }

            // Add columns.
            for (int i = 0; i < columns; i++)
            {
                table.Columns.Add();
            }

            // Add rows.
            foreach (var array in list)
            {
                table.Rows.Add(array);
            }

            return table;
        }
        */


        // WRITE IN EXCEL FILE

        //DataTable dataTable = new DataTable();
        //DataColumn column = new DataColumn("My Datacolumn");
        //DataColumn column2 = new DataColumn(@"C:\Developpement\Tests_dev");
        //dataTable.Columns.Add(column);
        //dataTable.Columns.Add(column2);
        //dataTable.Rows.Add(new object[] { "Foobar","BLABLA" });

        /*
        var lines = new List<string>();

        string[] columnNames = dt2.Columns.Cast<DataColumn>().
                                          Select(column => column.ColumnName).
                                          ToArray();

        var header = string.Join(",", columnNames);
        lines.Add(header);
        var valueLines = dt2.AsEnumerable()
                           .Select(row => string.Join(",", row.ItemArray));
        lines.AddRange(valueLines);
        File.WriteAllLines(Path.Combine(storePathFolder, "ExportExcel.csv"), lines);
        */



        /*
        // MessageBoxResult result = MessageBox.Show("Would you like to greet the world with a \"Hello, world\"?", "My App", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        MessageBoxResult result = MessageBox.Show("Are you sure to quite the application ?", strApplicationName, MessageBoxButton.YesNo, MessageBoxImage.Question); //, MessageBoxResult.No);
        switch (result)
        {
            case MessageBoxResult.Yes:
                this.Close();
                //MessageBox.Show("Hello to you too!", "My App");
                break;
            case MessageBoxResult.No:
                //MessageBox.Show("Oh well, too bad!", "My App");
                break;
           // case MessageBoxResult.Cancel:
           //     MessageBox.Show("Nevermind then...", "My App");
           //     break;
        }

        //this.Close();
        */

    }
}
