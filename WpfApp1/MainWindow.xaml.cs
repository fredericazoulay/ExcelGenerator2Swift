using System.Collections.Generic;
using System.Windows;

using System.Data;
using System.IO;
using ExcelDataReader;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using System.Reflection;
using System;
using System.ComponentModel;
using System.Linq;
using System.Configuration;

//using System.Configuration;
//using System.Collections.Specialized;

namespace ExcelGenerator2SwiftApp
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string strApplicationName = @"ExcelGenerator2Swift - v1.00";
        //string storePathFileExcelFileName = @"C:\Developpement\Tests_dev\TestSwift.xlsx";
        //string storePathFolderTextFiles = @"C:\Developpement\Tests_dev\";
        //string storeExtension = @".txt";
        string storePathFileExcelFileName = @"";
        string storePathFolderTextFiles = @"";
        string storeExtension = @"";
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();

        public MainWindow()
        {
            InitializeComponent();

            storePathFileExcelFileName = ConfigurationSettings.AppSettings.Get("ExcelFileName");
            storePathFolderTextFiles = ConfigurationSettings.AppSettings.Get("FolderTextFiles");
            storeExtension = ConfigurationSettings.AppSettings.Get("Extension");

            txbPathExcelFile.Text = storePathFileExcelFileName;
            txbPathDirectory.Text = storePathFolderTextFiles;
            txbExtension.Text = storeExtension;
        }



        #region Part1

        /// <summary>
        /// ExcelToDataTableUsingExcelDataReader
        /// </summary>
        /// <param name="storePath"></param>
        /// <returns></returns>
        public DataTable ExcelToDataTableUsingExcelDataReader(string storePath)
        {
            FileStream stream = File.Open(storePath, FileMode.Open, FileAccess.Read);

            string fileExtension = System.IO.Path.GetExtension(storePath);
            IExcelDataReader excelReader = null;
            if (fileExtension == ".xls")
            {
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (fileExtension == ".xlsx")
            {
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            //excelReader.IsFirstRowAsColumnNames = true;

            DataSet result = excelReader.AsDataSet();
            var test = result.Tables[0];

            stream.Close();
            return result.Tables[0];
        }

        /// <summary>
        /// btnBrowseExcelFile_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowseExcelFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = txbPathExcelFile.Text;
            if (openFileDialog.ShowDialog() == true)
            {
                txbPathExcelFile.Text = openFileDialog.FileName;
                //txbPath.Text = File.ReadAllText(openFileDialog.FileName);
            }
            storePathFileExcelFileName = txbPathExcelFile.Text;
        }

        /// <summary>
        /// btnRead_1_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRead_1_Click(object sender, RoutedEventArgs e)
        {
            dt1 = ExcelToDataTableUsingExcelDataReader(storePathFileExcelFileName);
            dataGrid_1.ItemsSource = dt1.AsDataView();
        }

        /// <summary>
        /// btnGenerate_1_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGenerate_1_Click(object sender, RoutedEventArgs e)
        {
            string strFileToWrite = "";

            int indexCol = 0;
            int indexRow = 0;
            foreach (DataColumn col in dt1.Columns)
            {
                List<string> lstLinesToWrite = new List<string>();
                foreach (DataRow row in dt1.Rows)
                {
                    if (indexRow == 0)
                    {
                        strFileToWrite = row.ItemArray[indexCol].ToString();
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(row.ItemArray[indexCol].ToString()))
                            lstLinesToWrite.Add(row.ItemArray[indexCol].ToString());
                    }
                    indexRow++;
                }
                indexRow = 0;
                indexCol++;
                System.IO.File.WriteAllLines(@strFileToWrite, lstLinesToWrite);
            }
            MessageBox.Show("Text file(s) generation done.", strApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
        }
        #endregion

        #region Part2

        /// <summary>
        /// btnBrowsePath_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowsePath_Click(object sender, RoutedEventArgs e)
        {
            var openFolderDialog = new VistaFolderBrowserDialog();
            openFolderDialog.SelectedPath = txbPathDirectory.Text;
            if (openFolderDialog.ShowDialog() == true)
            {
                txbPathDirectory.Text = openFolderDialog.SelectedPath;
            }
            storePathFolderTextFiles = txbPathDirectory.Text;
        }

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRead_2_Click(object sender, RoutedEventArgs e)
        {
            // Lectures des fichiers texte dans DataTable + DataGrid
            string[] arrFileEntries = Directory.GetFiles(storePathFolderTextFiles);
            List<string> lstFileEntries = new List<string>();
            foreach (string fileName in arrFileEntries)
            {
                
                //if (fileName.EndsWith(".txt"))
                if (fileName.EndsWith(txbExtension.Text))
                {
                    lstFileEntries.Add(fileName);
                }
            }

            List<string> lstFileRead = new List<string>();
            foreach (string fileName in lstFileEntries)
            {
                //string[] arrLinesToRead = System.IO.File.ReadAllLines(fileName);
                string linesToRead = System.IO.File.ReadAllText(fileName);
                //string strText = fileName + "\r\n" + linesToRead;
                string strText = linesToRead;
                lstFileRead.Add(strText);
            }

            DataTable outputTable = new DataTable();
            int index = 0;
            foreach (var newColumnName in lstFileEntries)
            {
                outputTable.Columns.Add("Colomn"+index, typeof(string));
                index++;
            }

            //outputTable.Columns.Add("Colomn0", typeof(string));
            //outputTable.Columns.Add("Colomn1", typeof(string));

            //foreach (var newColumnName in lstFileEntries)
            //{
            //    outputTable.Columns.Add(@newColumnName, typeof(string));
            //}

            outputTable.Rows.Add(lstFileEntries.ToArray());
            outputTable.Rows.Add(lstFileRead.ToArray());

            //foreach (var newColumnName in lstFileRead)
            //{
            //    //outputTable.Columns.Add(@newColumnName, typeof(string));
            //    outputTable.Rows.Add(lstFileRead.ToArray());
            //    //outputTable.Rows.Add(new object[] { "newColumnName", "BLAL" });
            //}

            /*
            //foreach (DataColumn inputColumn in inputTable.Columns)
            foreach (var inputColumnName in lstFileEntries)
            {
                //For each old column we generate a row in the new table 
                DataRow newRow = outputTable.NewRow();
                //newRow[inputColumnName] = "Hello";

                //Looks in the former header row to fill in the first column 
                //newRow[0] = inputColumnName.ToString();

                int counter = 0;
                //foreach (DataRow row in inputTable.Rows)
                foreach (string row in lstFileRead)
                {
                    //newRow[counter] = row[inputColumnName].ToString();
                    newRow[0] = row;
                    counter++;
                }
                
                outputTable.Rows.Add(newRow);
            }
            */

            /*
            foreach (string fileName in arrFileEntries)
            {
                lstList.Add(fileEntries);
                //ProcessFile(fileName);
                if (fileName.EndsWith(".txt"))
                {
                    List<string> lstLinesToRead = new List<string>();
                    string[] arrLinesToRead = System.IO.File.ReadAllLines(fileName);
                    foreach (string line in arrLinesToRead)
                    {
                        lstLinesToRead.Add(line);
                        //lstLinesToRead.Add(new string[] { line });
                    }
                    lstList.Add(lstLinesToRead,"");
                }
            }
            */

            //List<string[]> list = new List<string[]>();
            //list.Add(new string[] { "Column 1", "Column 2", "Column 3" });
            //list.Add(new string[] { "Row 2", "Row 2" });
            //list.Add(new string[] { "Row 3" });

            //dt2 = HelperConvert.ToDataTable(lstLinesToRead);
            //dt2 = ConvertListToDataTable(lstLinesToRead);
            dt2 = outputTable; // ConvertListToDataTable(lstFileRead);
            //dt2 = ExcelToDataTableUsingExcelDataReader(storePathFile);
            dataGrid_2.ItemsSource = outputTable.AsDataView();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGenerate_2_Click(object sender, RoutedEventArgs e)
        {
            // Ecriture fichier Excel à partir d'une DataTable



            HelperConvert.ExportExcelFile(dt2, Path.Combine(storePathFolderTextFiles, "ExportExcel.xlsx"));
            MessageBox.Show("Excel file generation done.", strApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);

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
        }

        #endregion

        
        private void OnClickAbout(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("About : " + strApplicationName, strApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
        }
        /// <summary>
        /// btnClose_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Are you sure to quite the application ?", strApplicationName, MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                this.Close();

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



}



