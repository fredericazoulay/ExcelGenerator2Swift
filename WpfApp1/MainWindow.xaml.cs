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

//****************************************************************************************************************
//****************************************************************************************************************
//****************************************************************************************************************
//****************************************************************************************************************

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
        string storePathFolderSwift_0 = @"";
        string storeSuffix_0 = @"";
        

        DataTable dt1 = new DataTable();
        DataTable dt2 = new DataTable();

        /// <summary>
        /// MainWindow
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            storePathFileExcelFileName = ConfigurationSettings.AppSettings.Get("ExcelFileName");
            storePathFolderTextFiles = ConfigurationSettings.AppSettings.Get("FolderTextFiles");
            storeExtension = ConfigurationSettings.AppSettings.Get("Extension");
            storePathFolderSwift = ConfigurationSettings.AppSettings.Get("FolderSwift");
            storeSuffix_0 = ConfigurationSettings.AppSettings.Get("Suffix_0");
            storePathFolderSwift_0 = ConfigurationSettings.AppSettings.Get("FolderSwift_0");

            txbPathExcelFile.Text = storePathFileExcelFileName;
            txbPathDirectory.Text = storePathFolderTextFiles;
            txbExtension.Text = storeExtension;
            txbPathSwift.Text = storePathFolderSwift;
            txbSuffix_0.Text = storeSuffix_0;
            txbPathSwift_0.Text = storePathFolderSwift_0;
        }

        #region Part0

        /// <summary>
        /// btnBrowse_0_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowse_0_Click(object sender, RoutedEventArgs e)
        {
            var openFolderDialog = new VistaFolderBrowserDialog();
            openFolderDialog.SelectedPath = txbPathSwift_0.Text;
            if (openFolderDialog.ShowDialog() == true)
            {
                txbPathSwift_0.Text = openFolderDialog.SelectedPath;
            }
            storePathFolderSwift_0 = txbPathSwift_0.Text;
        }

        /// <summary>
        /// btnDuplicate_0_BIG_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDuplicate_0_BIG_Click(object sender, RoutedEventArgs e)
        {
            var suffix = txbSuffix_0.Text;
            var pathFolderSwift = txbPathSwift_0.Text; // storePathFolderSwift_0

            string[] arrFileEntries = Directory.GetFiles(pathFolderSwift);


            List<string> lstFileEntries = new List<string>();
            foreach (string fullFileName_in in arrFileEntries)
            {
                FileInfo file = new FileInfo(fullFileName_in);
                string pathFile_in = file.DirectoryName;
                string pathFile_out = @pathFile_in + suffix;
                //string pathFile_out = Path.Combine(@pathFile_in, @suffix);
                string nameFile = file.Name;
                //string fullFileName_out = pathFile_out + nameFile;
                string fullFileName_out = Path.Combine(@pathFile_out, @nameFile);



                if (!Directory.Exists(@pathFile_out))
                {
                    DirectoryInfo di = Directory.CreateDirectory(@pathFile_out);
                }

                //string linesToRead = System.IO.File.ReadAllText(fullFileName_in);
                string[] linesToRead = System.IO.File.ReadAllLines(fullFileName_in);
                // Lines to write
                List<string> lstlinesToWrite = new List<string>();
                // Params
                string strStart = "{2:O";
                int lenStart = strStart.Length;
                string strTextEndLine = "\n";
                string typeSwift = "";
                string refFrontSwift = "";
                string strTextToFind = "";
                string[] arrTextToFind = { @":20:", @":20C::SEME//" };
                string strTextToFind_1 = @":20:";
                string strTextToFind_2 = @":20C::SEME//";

                foreach (string lineToRead in linesToRead)
                {
                    // DEBUT HEADER SWIFT
                    if (lineToRead.Contains(strStart))
                    {
                        int indexStartTypeSwift = lineToRead.IndexOf(@strStart);
                        int indexStart = lineToRead.IndexOf("{2:O");
                        typeSwift = lineToRead.Substring(indexStartTypeSwift + lenStart, 3);
                    }

                    // ONLY RefFront NEWM => No CANC
                    if (typeSwift == "304")
                    {
                        // :20:
                        strTextToFind = @":20:";
                    }
                    else if (typeSwift == "502")
                    {
                        // :20C::SEME//
                        strTextToFind = @":20C::SEME//";
                    }
                    else if ((typeSwift == "541") || (typeSwift == "543"))
                    {
                        // :20C::SEME//
                        strTextToFind = @":20C::SEME//";
                    }
                    else if (typeSwift == "598")
                    {
                        // :20:
                        strTextToFind = @":20:";
                    }
                    else
                    {
                        // Nothing or ADD other type of Swift or Replace by several strings
                    }


                    //******************************************
                    // Faire boucle si on veut ajouter d'autres valeurs de recherche, cela sera plus propre
                    //******************************************
                    if (lineToRead.StartsWith(strTextToFind_1)) // StartsWith or Contains ?
                    {
                        int indexStart1 = lineToRead.IndexOf(strTextToFind_1);
                        int indexEnd1 = lineToRead.Length;
                        if ((indexStart1 >= 0) && (indexEnd1 >= 0))
                        {
                            int indexLen = indexEnd1 + indexStart1 - strTextToFind_1.Length;
                            refFrontSwift = lineToRead.Substring(indexStart1 + strTextToFind_1.Length, indexLen);
                        }
                    }
                    if  ( lineToRead.StartsWith(strTextToFind_2) )
                    {
                        int indexStart1 = lineToRead.IndexOf(strTextToFind_2);
                        int indexEnd1 = lineToRead.Length;
                        if ((indexStart1 >= 0) && (indexEnd1 >= 0))
                        {
                            int indexLen = indexEnd1 + indexStart1 - strTextToFind_2.Length;
                            refFrontSwift = lineToRead.Substring(indexStart1 + strTextToFind_2.Length, indexLen);
                        }
                    }
                    //******************************************

                    if (refFrontSwift.Length > 0)
                    {
                        string refFrontSwiftUpdated = refFrontSwift + suffix;
                        string lineUpdated = lineToRead.Replace(refFrontSwift, refFrontSwiftUpdated);
                        lstlinesToWrite.Add(@lineUpdated);
                    }
                    else
                    {
                        lstlinesToWrite.Add(@lineToRead);
                    }

                }

                StreamWriter fileStream = System.IO.File.CreateText(@fullFileName_out);
                int indexLine = 1;
                foreach (string lineToWrite in lstlinesToWrite)
                {
                    if (indexLine < lstlinesToWrite.Count)
                    { 
                        fileStream.Write(lineToWrite + strTextEndLine); // Saut de ligne UNIX
                    }
                    else
                    { 
                        fileStream.Write(lineToWrite);
                    }
                    indexLine++;
                }
                fileStream.Close();


                lstFileEntries.Add(@fullFileName_out);
            }
            MessageBox.Show("Swift file(s) duplication done.", strApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
        }


        /// <summary>
        /// btnDuplicate_0_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDuplicate_0_Click(object sender, RoutedEventArgs e)
        {
            var suffix = txbSuffix_0.Text;
            var pathFolderSwift = txbPathSwift_0.Text; // storePathFolderSwift_0

            string[] arrFileEntries = Directory.GetFiles(pathFolderSwift);
            

            List<string> lstFileEntries = new List<string>();
            foreach (string fullFileName_in in arrFileEntries)
            {
                FileInfo file = new FileInfo(fullFileName_in);
                string pathFile_in = file.DirectoryName;
                string pathFile_out = @pathFile_in + suffix;
                //string pathFile_out = Path.Combine(@pathFile_in, @suffix);
                string nameFile = file.Name;
                //string fullFileName_out = pathFile_out + nameFile;
                string fullFileName_out = Path.Combine(@pathFile_out, @nameFile);



                if (!Directory.Exists(@pathFile_out))
                {
                    DirectoryInfo di = Directory.CreateDirectory(@pathFile_out);
                }

                string linesToRead = System.IO.File.ReadAllText(fullFileName_in);
                string strStart = "{2:O";
                int lenStart = strStart.Length;
                int indexStartTypeSwift = linesToRead.IndexOf(@strStart);

                int indexStart = linesToRead.IndexOf("{2:O");
                string typeSwift = linesToRead.Substring(indexStartTypeSwift + lenStart, 3);
                string refFrontSwift = "";
                string strTextToFind = "";

                // ONLY RefFront NEWM => No CANC
                if (typeSwift == "304")
                {
                    // :20:
                    strTextToFind = @":20:"; 
                }
                else if (typeSwift == "502")
                {
                    // :20C::SEME//
                    strTextToFind = @":20C::SEME//"; 
                }
                else if ((typeSwift == "541") || (typeSwift == "543"))
                {
                    // :20C::SEME//
                    strTextToFind = @":20C::SEME//";
                }
                else if (typeSwift == "598")
                {
                    // :20:
                    strTextToFind = @":20:";
                }
                else
                {
                    // Nothing
                }

                string strTextEndLine = "\n";
                int indexStart1 = linesToRead.IndexOf(@strTextToFind);
                int indexEnd1 = linesToRead.IndexOf(strTextEndLine, indexStart1);
                if ((indexStart1 > 0) && (indexEnd1 > 0))
                {
                    int indexLen1 = indexEnd1 - indexStart1;
                    refFrontSwift = linesToRead.Substring(indexStart1, indexLen1);
                }

                if (refFrontSwift.Length > 0)
                { 
                    string refFrontSwiftUpdated = refFrontSwift + suffix;

                    string linesUpdated = linesToRead.Replace(refFrontSwift, refFrontSwiftUpdated);
                    StreamWriter fileStream = System.IO.File.CreateText(@fullFileName_out);
                    fileStream.WriteLine(linesUpdated);
                    fileStream.Close();
                }

                lstFileEntries.Add(@fullFileName_out);
            }
            MessageBox.Show("Swift file(s) duplication done.", strApplicationName, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion


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
        /// btnRead_2_Click
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

        /// <summary>
        /// btnBrowse_3_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// generateSwift
        /// </summary>
        /// <param name="strSwiftType"></param>
        private void generateSwift(string strSwiftType)
        {
            int nbTestFiles = Int32.Parse(txbNbreTestFiles.Text);
            string fileNameIn = Path.Combine(txbPathSwift.Text, @"Swift\In\" + strSwiftType + ".txt");
            string PathOut = Path.Combine(txbPathSwift.Text, @"Swift\Out\" + strSwiftType);
            string fileNameOut = Path.Combine(txbPathSwift.Text, @"Swift\Out\" + strSwiftType + @"\ReceivedFiles\" + strSwiftType + "_");
            string linesToRead = System.IO.File.ReadAllText(fileNameIn);

            if (!Directory.Exists(@PathOut))
            {
                DirectoryInfo di = Directory.CreateDirectory(@PathOut);
            }

            for (int iCpt = 0; iCpt < nbTestFiles; iCpt++)
            {
                string linesIncremented = linesToRead.Replace("{CPT}", iCpt.ToString());
                string strFileNameOut = @fileNameOut + iCpt.ToString() + ".txt";
                StreamWriter fileStream = System.IO.File.CreateText(@strFileNameOut);
                fileStream.WriteLine(linesIncremented);
                fileStream.Close();
            }
        }

        /// <summary>
        /// btnMT502_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMT502_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT502");
        }

        /// <summary>
        /// btnMT54X_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMT54X_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT54X");
        }

        /// <summary>
        /// btnMT598_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMT598_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT598");
        }

        /// <summary>
        /// btnMT304_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMT304_Click(object sender, RoutedEventArgs e)
        {
            this.generateSwift("MT304");
        }

        /// <summary>
        /// btnUBIX_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// OnClickAbout
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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



