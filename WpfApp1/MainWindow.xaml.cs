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
        string storePathFolderSwift = @"";
        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();

        public MainWindow()
        {
            InitializeComponent();

            storePathFileExcelFileName = ConfigurationSettings.AppSettings.Get("ExcelFileName");
            storePathFolderTextFiles = ConfigurationSettings.AppSettings.Get("FolderTextFiles");
            storeExtension = ConfigurationSettings.AppSettings.Get("Extension");
            storePathFolderSwift = ConfigurationSettings.AppSettings.Get("FolderSwift");

            txbPathExcelFile.Text = storePathFileExcelFileName;
            txbPathDirectory.Text = storePathFolderTextFiles;
            txbExtension.Text = storeExtension;
            txbPathSwift.Text = storePathFolderSwift;
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
            FileInfo fi = new FileInfo(@txbPathExcelFile.Text);

            openFileDialog.InitialDirectory = fi.DirectoryName;
            if (openFileDialog.ShowDialog() == true)
            {
                txbPathExcelFile.Text = openFileDialog.FileName;
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
                //System.IO.File.WriteAllLines(@strFileToWrite, lstLinesToWrite);
                StreamWriter fileStream = System.IO.File.CreateText(@strFileToWrite);
                int cpt = 0;
                foreach (String line in lstLinesToWrite)
                {
                    if (cpt == lstLinesToWrite.Count-1)
                        fileStream.Write(line);
                    else
                        fileStream.WriteLine(line);
                    cpt++;
                }
                fileStream.Close();
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
            openFolderDialog.SelectedPath = @txbPathDirectory.Text;
            if (openFolderDialog.ShowDialog() == true)
            {
                txbPathDirectory.Text = openFolderDialog.SelectedPath;
            }
            storePathFolderTextFiles = txbPathDirectory.Text;
        }


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
                if (fileName.EndsWith(txbExtension.Text) || fileName.Contains(txbExtension.Text))
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
                if (!string.IsNullOrEmpty(strText))
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
        }

        #endregion

        #region Part3

        private void btnBrowse_3_Click(object sender, RoutedEventArgs e)
        {
            var openFolderDialog = new VistaFolderBrowserDialog();
            openFolderDialog.SelectedPath = txbPathSwift.Text;
            if (openFolderDialog.ShowDialog() == true)
            {
                txbPathSwift.Text = openFolderDialog.SelectedPath;
            }
            storePathFolderSwift = txbPathSwift.Text;
        }

        private void generateSwift(string strSwiftType)
        {
            string fileNameIn = Path.Combine(txbPathSwift.Text, @"Swift\In\" + strSwiftType + ".txt");
            string PathOut = Path.Combine(txbPathSwift.Text, @"Swift\Out\" + strSwiftType);
            string fileNameOut = Path.Combine(txbPathSwift.Text, @"Swift\Out\" + strSwiftType + @"\" + strSwiftType + "_");
            string linesToRead = System.IO.File.ReadAllText(fileNameIn);

            if (!Directory.Exists(@PathOut))
            {
                DirectoryInfo di = Directory.CreateDirectory(@PathOut);
            }

            for (int iCpt = 0; iCpt < 1000; iCpt++)
            {
                string linesIncremented = linesToRead.Replace("{CPT}", iCpt.ToString());
                string strFileNameOut = @fileNameOut + iCpt.ToString() + ".txt";
                //FileStream fileStreamCreate = System.IO.File.Create(strFileNameOut);
                //fileStreamCreate.Close();
                StreamWriter fileStream = System.IO.File.CreateText(@strFileNameOut);
                fileStream.WriteLine(linesIncremented);
                fileStream.Close();
                //System.IO.File.WriteAllLines(strFileNameOut, arrLinesIncremented);
            }
        }

        private void btnMT502_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT502");
        }

        private void btnMT54X_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT54X");
        }

        private void btnMT598_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT598");
        }

        private void btnMT304_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT304");
        }

        private void btnUBIX_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("UBIX");
        }

        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGenerate_3_Click(object sender, RoutedEventArgs e)
        {
            // Ecriture des fichiers de test
            this.generateSwift("MT502");
            this.generateSwift("MT54X");
            this.generateSwift("MT598");
            this.generateSwift("MT304");
            this.generateSwift("UBIX");
            MessageBox.Show("Swift test files generation done.", strApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
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
        }

    }



}



