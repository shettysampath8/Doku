using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Windows;

namespace Replacer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string excelFilePath = string.Empty;
        private string textFilePath = string.Empty;
        private string outputFilePath = string.Empty;

        public void IsReplaceButtonEnabled()
        {
            if (excelFilePath != string.Empty && textFilePath != string.Empty)
                ReplaceButton.IsEnabled = true;
            else
                ReplaceButton.IsEnabled = false;
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                excelFilePath = openFileDialog.FileName;

            if (File.Exists(excelFilePath) && excelFilePath.EndsWith(".xlsx"))
            {
                ExcelFileNameTextBlock.Text = excelFilePath;
                IsReplaceButtonEnabled();
            }
            else
            {
                ExcelFileNameTextBlock.Text = "Import MST plan";
                excelFilePath = string.Empty;
                MessageBox.Show("Not valid excel file");
            }
        }

        private void BrowseTextFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                textFilePath = openFileDialog.FileName;

            if (File.Exists(textFilePath) && textFilePath.EndsWith(".txt"))
            {
                TextFileNameTextBlock.Text = textFilePath;
                IsReplaceButtonEnabled();
            }
            else
            {
                TextFileNameTextBlock.Text = "Import Master Session file";
                textFilePath = string.Empty;
                MessageBox.Show("Not valid text file");
            }
        }

        private void BrowseOutputFileButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
        }

        private void ReplaceButton_Click(object sender, RoutedEventArgs e)
        {
            var dictionary = ReadExcelData(excelFilePath);
            var textStrings = File.ReadAllLines(textFilePath);

            for(int i = 0; i < textStrings.Length; i++)
            {
                var currentLine = textStrings[i];

                var stringsInQuotes = new Regex("\".*?\"").Matches(currentLine);
                foreach(var str in stringsInQuotes)
                {
                    var strWithTrim = str.ToString().Trim('"');

                    if (dictionary.ContainsKey(strWithTrim))
                    {
                        currentLine = currentLine.Replace(strWithTrim, dictionary[strWithTrim]);
                    }
                }
                
                textStrings[i] = currentLine;       
            }


            var fileName = Path.GetFileName(textFilePath).Split('.')[0];
            var snapshotDirectory = Path.Combine(Directory.GetParent(textFilePath).FullName, "Snapshot.plmxml");
            if(Directory.Exists(snapshotDirectory))
            {
                MessageBox.Show($"A folder with name \"Snapshot.plmxml\"  already exists. please delete it and proceed further");
                Close();
            }

            var snapDirInfo  = Directory.CreateDirectory(snapshotDirectory);
            var outputFilePath = Path.Combine(snapshotDirectory, fileName + ".plmxml");
            File.WriteAllLines(outputFilePath, textStrings);
            string outputFileName = snapDirInfo.FullName + ".zip";
            //if(!Compress(snapDirInfo.FullName, outputFileName))
            //{
            //    MessageBox.Show($"There was a error while generating output. Please clean folder and try again");
            //    Close();
            //}

            try
            {
                ZipFile.CreateFromDirectory(snapDirInfo.FullName, outputFileName);
            }
            catch {

                MessageBox.Show($"There was a error while generating output. Please clean folder and try again");
                Close();
            }
            MessageBox.Show($"Output zip file generated successfully at {snapshotDirectory}");
            Close();
        }

        public static bool Compress(string sInDir, string sOutFile)
        {
            string[] sFiles = Directory.GetFiles(sInDir, "*.*", SearchOption.AllDirectories);
            int iDirLen = sInDir[sInDir.Length - 1] == Path.DirectorySeparatorChar ? sInDir.Length : sInDir.Length + 1;


            using (FileStream outFile = new FileStream(sOutFile, FileMode.Create, FileAccess.Write, FileShare.None))
            using (GZipStream str = new GZipStream(outFile, CompressionMode.Compress))
                foreach (string sFilePath in sFiles)
                {
                    string sRelativePath = sFilePath.Substring(iDirLen);
                    CompressFile(sInDir, sRelativePath, str);
                }
       
            return true;
        }

        static void CompressFile(string sDir, string sRelativePath, GZipStream zipStream)
        {
            //Compress file name
            char[] chars = sRelativePath.ToCharArray();
            zipStream.Write(BitConverter.GetBytes(chars.Length), 0, sizeof(int));
            foreach (char c in chars)
                zipStream.Write(BitConverter.GetBytes(c), 0, sizeof(char));

            //Compress file content
            byte[] bytes = File.ReadAllBytes(Path.Combine(sDir, sRelativePath));
            zipStream.Write(BitConverter.GetBytes(bytes.Length), 0, sizeof(int));
            zipStream.Write(bytes, 0, bytes.Length);
        }

        private Dictionary<string,string> ReadExcelData(string fileName)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            //List<Tuple<string, string>> tuples = new List<Tuple<string, string>>();

            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }


            var rowCount = dtexcel.Rows.Count;
            for(int i = 1; i < rowCount; i++)
            {
                var currentRow = dtexcel.Rows[i];
                var mainStirng = currentRow[0] as string;
                var replacableString = currentRow[1] as string;

                if (!string.IsNullOrEmpty(mainStirng) && !string.IsNullOrEmpty(replacableString))
                    dict.Add(mainStirng, replacableString);
                        //tuples.Add(new Tuple<string, string>(mainStirng, replacableString)) ;
            }

            return dict;
        }
    }
}
