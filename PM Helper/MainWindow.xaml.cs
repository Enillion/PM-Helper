using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
//using System.Windows.Data;
//using System.Windows.Documents;
//using System.Windows.Input;
//using System.Windows.Media;
//using System.Windows.Media.Imaging;
//using System.Windows.Navigation;
//using System.Windows.Shapes;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using Microsoft.Win32;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO.Compression;
//using HtmlAgilityPack;

namespace PM_Helper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IDictionary<string, string> languageDictionary = new Dictionary<string, string>
        {
            {"Albanian", "sq-al"},
            {"Arabic (Egypt)", "ar-eg"},
            {"Arabic (Saudi Arabia)", "ar-sa"},
            {"Arabic (United Arab Emirates)", "ar-sa"},
            {"Azeri (Latin - Azerbaijan)", "az"},
            {"Bosnian (Latin)", "bs"},
            {"Bulgarian", "bg-bg"},
            {"Burmese", "my"},
            {"Catalan", "ca-es"},
            {"Catalan (Spain)", "ca-es"},
            {"Chinese (Hong Kong)", "zh-hk"},
            {"Chinese (Simplified)", "zh-cn"},
            {"Chinese (Traditional)", "zh-tw"},
            {"Croatian", "hr-hr"},
            {"Czech", "cs-cz"},
            {"Danish", "da-dk"},
            {"Dari", "prs-af"},
            {"Dutch", "nl-nl"},
            {"Dutch (Belgium)", "nl-be"},
            {"English (Canada)", "en-ca"},
            {"English (US)", "en-us"},
            {"English (UK)", "en-gb"},
            {"Estonian", "et-ee"},
            {"Finnish", "fi-fi"},
            {"French (Canada)", "fr-ca"},
            {"French (France)", "fr-fr"},
            {"Georgian", "ka"},
            {"German (Germany)", "de-de"},
            {"Greek", "el-gr"},
            {"Hebrew", "he-il"},
            {"Hindi", "hi-in"},
            {"Hungarian", "hu-hu"},
            {"Icelandic", "is-is"},
            {"Indonesian", "id-id"},
            {"Irish (Gaelic)", "ga"},
            {"Italian", "it-it"},
            {"Japanese", "ja-jp"},
            {"Kannada", "kn"},
            {"Kazakh", "kk"},
            {"Khmer", "km"},
            {"Korean", "ko-kr"},
            {"Laothian", "lo"},
            {"Latvian", "lv-lv"},
            {"Lithuanian", "lt-lt"},
            {"Macedonian", "mk-mk"},
            {"Malay (Malaysia)", "ms-my"},
            {"Maltese", "mt"},
            {"Marathi", "mr"},
            {"Norwegian", "nn-no"},
            {"Norwegian Bokmaal", "nb-no"},
            {"Persian", "fa-ir"},
            {"Polish", "pl-pl"},
            {"Portuguese", "pt-pt"},
            {"Portuguese (Brazil)", "pt-br"},
            {"Romanian", "ro-ro"},
            {"Russian", "ru-ru"},
            {"Serbian - Serbia (Latin)", "srl-rs"},
            {"Serbian (Cyrillic)", "sr-rs"},
            {"Slovak", "sk-sk"},
            {"Slovenian", "sl-si"},
            {"Spanish (Colombia)", "es-co"},
            {"Spanish (Latin America)", "es-xl"},
            {"Spanish (Spain)", "es-es"},
            {"Swedish", "sv-se"},
            {"Tagalog", "tl"},
            {"Thai", "th-th"},
            {"Turkish", "tr-tr"},
            {"Ukrainian", "uk-ua"},
            {"Uzbek (Latin)", "uzl"},
            {"Vietnamese", "vi-vn"},
            {"Welsh", "cy"},
        };

        public MainWindow()
        {
            InitializeComponent();
        }

        private void StackDropSource_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] path = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (System.IO.Path.GetExtension(path[0]) == ".xlsx" || System.IO.Path.GetExtension(path[0]) == ".zip")
                {                    
                    XlsxPathDisplay.Text = path[0];
                }
                else
                {
                    MessageBox.Show("Incorrect file extension, only XLSX and ZIP files are supported.");
                }                
            }
        }

        private void BrowseFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                if (System.IO.Path.GetExtension(openFileDialog.FileName) == ".xlsx")
                {
                    XlsxPathDisplay.Text = openFileDialog.FileName;
                }
                else
                {
                    MessageBox.Show("Incorrect file extension, only XLSX files are supported.");
                }
            }
        }

        private void Convert_Click(object sender, RoutedEventArgs e)
        {            
            if (XlsxPathDisplay.Text == "")
            {
                MessageBox.Show("XTM generated XLSX mertic file was not selected.");
            }
            else
            {                
                string[] arguments = { XlsxPathDisplay.Text, (DropdownList.SelectedItem.ToString()).Substring(38, ((DropdownList.SelectedItem.ToString()).Length - 38)), CheckBox_100.IsChecked.Value.ToString() }; // Target language is the substring of value "System.Windows.Controls.Comboboxitem: <comboboxItem>"
                ProgressBar.Value = 0;
                ProgressBar.IsIndeterminate = true;
                ConvStack.IsEnabled = false; // Disable tab UI
                BackgroundWorker bgw = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw.WorkerReportsProgress = true;
                bgw.DoWork += Bgw_DoWork;
                bgw.ProgressChanged += Bgw_ProgressChanged;
                bgw.RunWorkerCompleted += Bgw_RunWorkerCompleted;
                bgw.RunWorkerAsync(arguments);
            }
        }

        private void Bgw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ConvStack.IsEnabled = true; // Enable tab UI
            ProgressBar.IsIndeterminate = false;
            MessageBox.Show("Done");
        }

        private void Bgw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
        }

        private void Bgw_DoWork(object sender, DoWorkEventArgs e)
        {            
            string[] arguments = (string[])e.Argument;
            string filePath = arguments[0];
            string sourceLanguage = arguments[1];
            bool ignore100matches = false;
            string excelPath = "";
            if (arguments[2] == "True")
            {
                ignore100matches = true;
            }
            else if (arguments[2] == "False")
            {
                ignore100matches = false;
            }

            try
            {
                if (System.IO.Path.GetExtension(filePath) == ".zip")
                {
                    ZipFile.ExtractToDirectory(filePath, filePath.Substring(0, (filePath.LastIndexOf(@"\"))));
                    ZipArchive zip = ZipFile.OpenRead(filePath);
                    if (zip.Entries.Count == 1)
                    {
                        foreach (ZipArchiveEntry entry in zip.Entries)
                        {
                            if (entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || entry.FullName.EndsWith(".XLSX", StringComparison.OrdinalIgnoreCase))
                            {
                                excelPath = filePath.Substring(0, (filePath.LastIndexOf(@"\") + 1)) + entry.FullName;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Incorrect Zip file! Provided archive consists of multiple files.");
                        return;
                    }
                }
                else
                {
                    excelPath = filePath;
                }
            }
            catch (Exception eex)
            {
                MessageBox.Show("Error! " + eex.Message + " Conversion Process terminated.");
                return;
            }
            
            try
            {
                ReadExcelFileAndCreateLog(excelPath, sourceLanguage, ignore100matches);                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void ReadExcelFileAndCreateLog(string excelPath, string sourceLanguage, bool ignore100matches)
        {
            //Create excel instance
            Excel.Application excel = new Excel.Application();
            //Create target workbook
            Excel.Workbook workbook = excel.Workbooks.Open(excelPath);
            //Open the source sheet
            Excel.Worksheet sourceSheet = workbook.Worksheets[1];
            //Get the used Range
            Excel.Range usedRange = sourceSheet.UsedRange;
            int range = usedRange.Rows.Count;
            string projectName = (string)(sourceSheet.Cells[2, 1] as Excel.Range).Value;

            List<string> languagesList = ExtractLanguages(range, sourceSheet);// extract list of languages from xlsx                
            foreach (string language in languagesList)
            {
                List<FileForTranslation> files = CreateListOfFiles(sourceSheet, excel, language, range, sourceLanguage);// create list of files translated into specified language
                string logLocationPath = excelPath.Substring(0, (excelPath.Length - (excelPath.Length - excelPath.LastIndexOf(@"\"))) + 1);
                string logName = projectName + "_" + (languageDictionary[language]).ToUpper();
                string logFilePath;
                if (ignore100matches)
                {
                    logFilePath = logLocationPath + logName + "_without_100%_matches.txt";
                }
                else
                {
                    logFilePath = logLocationPath + logName + "_with_100%_matches.txt";
                }
                // Create a new log file     
                CreateLogFile(files, projectName, sourceLanguage, logFilePath, language, ignore100matches);
            }

            //Close workbook                    
            workbook.Close(false);
            //Kill excelapp
            excel.Quit();
        }

        private void CreateLogFile(List<FileForTranslation> files, string projectName, string sourceLanguage, string logFilePath, string language, bool ignore100matches)
        {
            try
            {
                int totalTranslated = 0;
                int totalContextMatch = 0;
                int totalRepetitions = 0;
                int total100 = 0;
                int total95_99 = 0;
                int total85_94 = 0;
                int total75_84 = 0;
                int totalMachineTrans = 0;
                int totalNoMatch = 0;
                int totalTotal = 0;                
                // Check if file already exists. If yes, delete it.     
                if (File.Exists(logFilePath))
                {
                    File.Delete(logFilePath);
                }
                //Create file
                using (StreamWriter sw = File.CreateText(logFilePath))
                {
                    foreach (FileForTranslation file in files)
                    {
                        sw.WriteLine(@"File:               " + file.Name);                        
                        sw.WriteLine(@"Date:               " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss CET"));
                        sw.WriteLine(@"Project:            " + projectName);
                        sw.WriteLine(@"Language direction: " + file.SourceLanguageCode + " > " + file.TargetLanguageCode);
                        sw.WriteLine(@"");
                        sw.WriteLine(@"Match Types   Segments        Words      Percent");
                        sw.WriteLine(@"Translated           0" + ((new string(' ', (13 - file.NonTranslatable.ToString().Length))) + file.NonTranslatable.ToString()) + "            0");
                        sw.WriteLine(@"Context Match        0" + ((new string(' ', (13 - file.IceMatch.ToString().Length))) + file.IceMatch.ToString()) + "            0");
                        sw.WriteLine(@"Repetitions          0" + ((new string(' ', (13 - file.Repetitions.ToString().Length))) + file.Repetitions.ToString()) + "            0");
                        sw.WriteLine(@"Format Change        0            0            0");                        
                        sw.WriteLine(@"100%                 0" + ((new string(' ', (13 - file.Match100.ToString().Length))) + file.Match100.ToString()) + "            0");
                        sw.WriteLine(@"95% - 99%            0" + ((new string(' ', (13 - file.Match95_99.ToString().Length))) + file.Match95_99.ToString()) + "            0");
                        sw.WriteLine(@"85% - 94%            0" + ((new string(' ', (13 - file.Match85_94.ToString().Length))) + file.Match85_94.ToString()) + "            0");
                        sw.WriteLine(@"75% - 84%            0" + ((new string(' ', (13 - file.Match75_84.ToString().Length))) + file.Match75_84.ToString()) + "            0");
                        sw.WriteLine(@"50% - 74%            0            0            0");
                        sw.WriteLine(@"Machine Trans        0" + ((new string(' ', (13 - file.MachineTrans.ToString().Length))) + file.MachineTrans.ToString()) + "            0");
                        sw.WriteLine(@"No Match             0" + ((new string(' ', (13 - file.NoMatch.ToString().Length))) + file.NoMatch.ToString()) + "            0");
                        int total;
                        if (ignore100matches == true)
                        {
                            total = file.TotalWithout100();
                        }
                        else
                        {
                            total = file.TotalIncluding100();
                        }
                        sw.WriteLine(@"Total                0" + ((new string(' ', (13 - total.ToString().Length))) + total.ToString()) + "            0");
                        sw.WriteLine(@"Chars/word           0");
                        sw.WriteLine(@"");
                        sw.WriteLine(@"");
                        totalTranslated += file.NonTranslatable;
                        totalContextMatch += file.IceMatch;
                        totalRepetitions += file.Repetitions;
                        total100 += file.Match100;
                        total95_99 += file.Match95_99;
                        total85_94 += file.Match85_94;
                        total75_84 += file.Match75_84;
                        totalMachineTrans += file.MachineTrans;
                        totalNoMatch += file.NoMatch;
                        totalTotal += total;
                    }

                    sw.WriteLine(@"Total:" + ((new string(' ', (15 - files.Count.ToString().Length))) + files.Count.ToString()) + " files");
                    sw.WriteLine(@"Date:               " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss CET"));
                    sw.WriteLine(@"Project:            " + projectName);
                    sw.WriteLine(@"Language direction: " + languageDictionary[sourceLanguage] + " > " + languageDictionary[language]);
                    sw.WriteLine(@"");
                    sw.WriteLine(@"Match Types   Segments        Words      Percent");
                    sw.WriteLine(@"Translated           0" + ((new string(' ', (13 - totalTranslated.ToString().Length))) + totalTranslated.ToString()) + "            0");
                    sw.WriteLine(@"Context Match        0" + ((new string(' ', (13 - totalContextMatch.ToString().Length))) + totalContextMatch.ToString()) + "            0");
                    sw.WriteLine(@"Repetitions          0" + ((new string(' ', (13 - totalRepetitions.ToString().Length))) + totalRepetitions.ToString()) + "            0");
                    sw.WriteLine(@"Format Change        0            0            0");
                    sw.WriteLine(@"100%                 0" + ((new string(' ', (13 - total100.ToString().Length))) + total100.ToString()) + "            0");
                    sw.WriteLine(@"95% - 99%            0" + ((new string(' ', (13 - total95_99.ToString().Length))) + total95_99.ToString()) + "            0");
                    sw.WriteLine(@"85% - 94%            0" + ((new string(' ', (13 - total85_94.ToString().Length))) + total85_94.ToString()) + "            0");
                    sw.WriteLine(@"75% - 84%            0" + ((new string(' ', (13 - total75_84.ToString().Length))) + total75_84.ToString()) + "            0");
                    sw.WriteLine(@"50% - 74%            0            0            0");
                    sw.WriteLine(@"Machine Trans        0" + ((new string(' ', (13 - totalMachineTrans.ToString().Length))) + totalMachineTrans.ToString()) + "            0");
                    sw.WriteLine(@"No Match             0" + ((new string(' ', (13 - totalNoMatch.ToString().Length))) + totalNoMatch.ToString()) + "            0");
                    sw.WriteLine(@"Total                0" + ((new string(' ', (13 - totalTotal.ToString().Length))) + totalTotal.ToString()) + "            0");
                    sw.WriteLine(@"Chars/word           0");
                    sw.WriteLine(@"");
                    sw.WriteLine(@"");
                }
            }
            catch (Exception Exe)
            {
                MessageBox.Show(Exe.Message + "\n" + Exe.StackTrace);
            }

            
        }

        private List<FileForTranslation> CreateListOfFiles(Excel.Worksheet sourceSheet, Excel.Application excel, string language, int range, string sourceLanguage)
        {
            List<FileForTranslation> files = new List<FileForTranslation>();// create empty list of files translated into specific language
            for (int i = 3; i <= range; i++)
            {
                if (((string)(sourceSheet.Cells[i, 2] as Excel.Range).Value) == language && ((string)(sourceSheet.Cells[i, 1] as Excel.Range).Value) != "All")
                {
                    FileForTranslation file = new FileForTranslation(
                        (string)(sourceSheet.Cells[i, 1] as Excel.Range).Value,
                        sourceLanguage,
                        (string)(sourceSheet.Cells[i, 2] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 11] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 6] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 7] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 8] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 9] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 10] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 15] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 4] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 5] as Excel.Range).Value,
                        (int)(sourceSheet.Cells[i, 3] as Excel.Range).Value
                        );
                    file.SourceLanguageCode = languageDictionary[file.SourceLanguage];
                    file.TargetLanguageCode = languageDictionary[file.TargetLanguage];
                    files.Add(file);// populate previously crated list                    
                }
            }
            return files;
        }

        private List<string> ExtractLanguages(int range, Excel.Worksheet sourceSheet)
        {
            List<string> languages = new List<string>();
            for (int i = 3; i <= range; i++)
            {
                string language = (string)(sourceSheet.Cells[i, 2] as Excel.Range).Value;
                bool alreadyOnTheList = false;
                foreach (string lang in languages)
                {
                    if (lang == language)
                    {
                        alreadyOnTheList = true;
                    }
                }
                if (alreadyOnTheList == false)
                {
                    languages.Add(language);
                }
            }
            return languages;
        }

        //Metric Merger
        private void StackDropTarget_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] path = (string[])e.Data.GetData(DataFormats.FileDrop);
                FileAttributes attr = File.GetAttributes(path[0]);

                if (attr.HasFlag(FileAttributes.Directory))
                {
                    XlsxTargetPathDisplay.Text = path[0];
                }
                else
                {
                    MessageBox.Show("Please drag and drop folder containing XTM generated metrics");
                }
            }
        }

        private void Merge_Click(object sender, RoutedEventArgs e)
        {
            if (XlsxTargetPathDisplay.Text == "")
            {
                MessageBox.Show("Folder with XTM generated mertic files was not selected.");
            }
            else
            {
                string[] arguments = { XlsxTargetPathDisplay.Text, (DropdownList2.SelectedItem.ToString()).Substring(38, ((DropdownList2.SelectedItem.ToString()).Length - 38)), CheckBox_100_2.IsChecked.Value.ToString() };
                ProgressBar.Value = 0;
                ProgressBar.IsIndeterminate = true;
                MergeStack.IsEnabled = false; // Disable tab UI
                BackgroundWorker bgw = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw.WorkerReportsProgress = true;
                bgw.DoWork += Bgw_DoWork1;
                bgw.ProgressChanged += Bgw_ProgressChanged1;
                bgw.RunWorkerCompleted += Bgw_RunWorkerCompleted1;
                bgw.RunWorkerAsync(arguments);
            }
        }

        private void Bgw_RunWorkerCompleted1(object sender, RunWorkerCompletedEventArgs e)
        {
            MergeStack.IsEnabled = true; // Enable tab UI
            ProgressBar.IsIndeterminate = false;
            MessageBox.Show("Done");
        }

        private void Bgw_ProgressChanged1(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
        }

        private void Bgw_DoWork1(object sender, DoWorkEventArgs e)
        {
            string[] arguments = (string[])e.Argument;
            string folderPath = arguments[0];
            string sourceLanguage = arguments[1];
            bool ignore100matches = false;            
            if (arguments[2] == "True")
            {
                ignore100matches = true;
            }
            else if (arguments[2] == "False")
            {
                ignore100matches = false;
            }
            List<string> tempFilesList = new List<string>();            
            try
            {
                string[] defaultFilePaths = Directory.GetFiles(folderPath);
                foreach (string filePath in defaultFilePaths)
                {
                    if (System.IO.Path.GetExtension(filePath) == ".zip")
                    {
                        ZipFile.ExtractToDirectory(filePath, filePath.Substring(0, (filePath.LastIndexOf(@"\"))));
                        File.Delete(filePath);
                    }
                }
                string[] finalFilePaths = Directory.GetFiles(folderPath);
                string targetPath = finalFilePaths[0];
                for (int i = 1; i < finalFilePaths.Length; i++)
                {
                    string mergedFilePath = targetPath.Substring(0, (targetPath.Length - (targetPath.Length - targetPath.LastIndexOf(@"\")))) + @"\MERGED_" + (i - 1) + @".xlsx";
                    string previousMergedFilePath = targetPath.Substring(0, (targetPath.Length - (targetPath.Length - targetPath.LastIndexOf(@"\")))) + @"\MERGED_" + (i - 2) + @".xlsx";
                    tempFilesList.Add(mergedFilePath);
                    if (File.Exists(previousMergedFilePath))
                    {
                        Dictionary<string, List<FileForTranslation>> sourceContent = GetMetricContent(finalFilePaths[i]);
                        UpdateTarget(previousMergedFilePath, sourceContent, (i - 1));
                    }
                    else
                    {
                        Dictionary<string, List<FileForTranslation>> sourceContent = GetMetricContent(finalFilePaths[i]);
                        UpdateTarget(targetPath, sourceContent, (i - 1));
                    }                    
                }
                string finalMergedFile = targetPath.Substring(0, (targetPath.Length - (targetPath.Length - targetPath.LastIndexOf(@"\")))) + @"\MERGED.xlsx";
                File.Move(tempFilesList.Last(), finalMergedFile);
                foreach (string file in tempFilesList)
                {
                    File.Delete(file);
                }

                ReadExcelFileAndCreateLog(finalMergedFile, sourceLanguage, ignore100matches);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void UpdateTarget(string targetPath, Dictionary<string, List<FileForTranslation>> sourceContent, int counter)
        {
            string finalPath = targetPath.Substring(0, (targetPath.Length - (targetPath.Length - targetPath.LastIndexOf(@"\")))) + @"\MERGED_" + counter + @".xlsx";
            // Check if file already exists. If yes, delete it.     
            if (File.Exists(finalPath))
            {
                File.Delete(finalPath);
            }            

            //Create excel instance
            Excel.Application excel = new Excel.Application();
            //Create target workbook
            Excel.Workbook workbook = excel.Workbooks.Open(targetPath);
            //Open the source sheet
            Excel.Worksheet targetSheet = workbook.Worksheets[1];
            //Get the used Range
            Excel.Range usedRange = targetSheet.UsedRange;
            int range = usedRange.Rows.Count;
            int excelIndex = range;
            while ((string)(targetSheet.Cells[excelIndex, 1] as Excel.Range).Value == "All" || (string)(targetSheet.Cells[excelIndex, 1] as Excel.Range).Value == "")
            {
                Excel.Range TempRange = targetSheet.Cells.Range[targetSheet.Cells[excelIndex, 1], targetSheet.Cells[excelIndex, 158]];
                TempRange.Cells.Clear();                
                excelIndex--;
            }
            excelIndex += 1;
            //targetSheet.Cells[excelIndex + 1, 1] = "Row = " + (excelIndex + 1); 
            foreach (var lang in sourceContent)
            {
                string language = (lang.ToString()).Substring(1, ((lang.ToString()).Length - ((lang.ToString()).Length - (lang.ToString()).IndexOf(',')) - 1));                
                List<FileForTranslation> files = sourceContent[language];
                foreach (FileForTranslation file in files)
                {
                    targetSheet.Cells[excelIndex, 1] = file.Name;
                    targetSheet.Cells[excelIndex, 2] = file.TargetLanguage;
                    targetSheet.Cells[excelIndex, 11] = file.Repetitions;
                    targetSheet.Cells[excelIndex, 6] = file.Match100;
                    targetSheet.Cells[excelIndex, 7] = file.Match95_99;
                    targetSheet.Cells[excelIndex, 8] = file.Match85_94;
                    targetSheet.Cells[excelIndex, 9] = file.Match75_84;
                    targetSheet.Cells[excelIndex, 10] = file.MachineTrans;
                    targetSheet.Cells[excelIndex, 15] = file.NoMatch;
                    targetSheet.Cells[excelIndex, 4] = file.NonTranslatable;
                    targetSheet.Cells[excelIndex, 5] = file.IceMatch;
                    targetSheet.Cells[excelIndex, 3] = file.XTMtotal;
                    excelIndex += 1;
                }
            }            
            targetSheet.Cells[2, 11] = "N/A";
            targetSheet.Cells[2, 6] = "N/A";
            targetSheet.Cells[2, 7] = "N/A";
            targetSheet.Cells[2, 8] = "N/A";
            targetSheet.Cells[2, 9] = "N/A";
            targetSheet.Cells[2, 10] = "N/A";
            targetSheet.Cells[2, 15] = "N/A";
            targetSheet.Cells[2, 4] = "N/A";
            targetSheet.Cells[2, 5] = "N/A";
            targetSheet.Cells[2, 3] = "N/A";            

            //Save work book (.xlsx format)
            workbook.SaveAs(finalPath, Excel.XlFileFormat.xlOpenXMLWorkbook, null, null, false,
            false, Excel.XlSaveAsAccessMode.xlShared, false, false, null, null, null);
            //Close workbook                    
            workbook.Close(false);
            //Kill excelapp
            excel.Quit();

            //Excel.Range TempRange = ExcelWorksheet.get_Range("H11", "J15");
            //string language = (string)(sourceSheet.Cells[i, 2] as Excel.Range).Value;
        }

        private Dictionary<string, List<FileForTranslation>> GetMetricContent(string sourcePath)
        {
            //Create excel instance
            Excel.Application excel = new Excel.Application();
            //Create target workbook
            Excel.Workbook workbook = excel.Workbooks.Open(sourcePath);
            //Open the source sheet
            Excel.Worksheet sourceSheet = workbook.Worksheets[1];
            //Get the used Range
            Excel.Range usedRange = sourceSheet.UsedRange;
            int range = usedRange.Rows.Count;
            Dictionary<string, List<FileForTranslation>> SourceContent = new Dictionary<string, List<FileForTranslation>>();
            List<string> languagesList = ExtractLanguages(range, sourceSheet);// extract list of languages from xlsx 
            foreach (string language in languagesList)
            {
                List<FileForTranslation> files = CreateListOfFiles(sourceSheet, excel, language, range, "English (US)");// create list of files translated into specified language
                SourceContent.Add(language, files);
            }
            //Close workbook                    
            workbook.Close(false);
            //Kill excelapp
            excel.Quit();
            return SourceContent;
        }

        // LBT EXTRACTOR

        private void HtmlDropTarget_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] path = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (System.IO.Path.GetExtension(path[0]) == ".xml")
                {
                    HtmlPathDisplay.Text = path[0];
                }
                else
                {
                    MessageBox.Show("Incorrect file extension, only .xml LTB log files are supported.");
                }
            }
        }

        private void BrowseLTBFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                if (System.IO.Path.GetExtension(openFileDialog.FileName) == ".xml")
                {
                    HtmlPathDisplay.Text = openFileDialog.FileName;
                }
                else
                {
                    MessageBox.Show("Incorrect file extension, only .xml LTB log files are supported.");
                }
            }
        }

        private void Extract_Click(object sender, RoutedEventArgs e)
        {
            if (HtmlPathDisplay.Text == "")
            {
                MessageBox.Show("LTB log file was not selected.");
            }
            else
            {
                string[] arguments = { HtmlPathDisplay.Text, CheckBox_Multi.IsChecked.Value.ToString() };
                ProgressBar.Value = 0;
                ProgressBar.IsIndeterminate = true;
                HtmlStack.IsEnabled = false; // Disable tab UI
                BackgroundWorker bgw2 = new BackgroundWorker(); // initialization of backgroundworker to move app processing to the 2nd thread
                bgw2.WorkerReportsProgress = true;
                bgw2.DoWork += Bgw2_DoWork;
                bgw2.ProgressChanged += Bgw2_ProgressChanged;
                bgw2.RunWorkerCompleted += Bgw2_RunWorkerCompleted;
                bgw2.RunWorkerAsync(arguments);
            }
        }

        private void Bgw2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            HtmlStack.IsEnabled = true; // Enable tab UI
            ProgressBar.IsIndeterminate = false;
            MessageBox.Show("Done");
        }

        private void Bgw2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
        }

        private void Bgw2_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] arguments = (string[])e.Argument;
            string filePath = arguments[0];
            bool multi = false;
            if (arguments[1] == "True")
            {
                multi = true;
            }
            else if (arguments[1] == "False")
            {
                multi = false;
            }

            try
            {
                ExtractLTB(filePath, multi);
            }
            catch (Exception eeex)
            {
                MessageBox.Show("Error! " + eeex.Message + " Extraction process terminated.");
                return;
            }
        }

        private void ExtractLTB(string file, bool multi)
        {
            if (System.IO.Path.GetExtension(file) == ".xml")
            {
                try
                {
                    string exportFilePath = file.Substring(0, (file.Length - (file.Length - file.LastIndexOf(@"\"))) + 1);
                    string[] files = Directory.GetFiles(exportFilePath, "*.xml");


                    foreach (string subFile in files)
                    {
                        XmlDocument xml = new XmlDocument();
                        using (XmlReader xr = new XmlTextReader(subFile) { Namespaces = false })
                        {
                            xml.Load(xr);
                        }

                        var root = xml.SelectNodes("//*");
                        int rowCounter = 0;

                        //string exportFile = exportFilePath + "LTB_Log_Export.txt"; - obsolete
                        //string exportFile = subFile.Substring(0, (subFile.Length - (subFile.Length - subFile.LastIndexOf(@"."))) + 1) + "xml";
                        string exportFile = subFile.Substring(0, (subFile.Length - (subFile.Length - subFile.LastIndexOf(@".")))) + "_EXPORT.xml";

                        if (File.Exists(exportFile))
                        {
                            File.Delete(exportFile);
                        }

                        if (multi == false)
                        {
                            using (var sw = new StreamWriter(exportFile, true))
                            {
                                sw.WriteLine("<?xml version=\"1.0\" ?>");
                                sw.WriteLine("<extract>");
                                foreach (var element in root)
                                {
                                    var node = (XmlElement)element;
                                    if (node.Name == "Row" && !(node.ChildNodes[4] == null) || node.Name == "ss:Row" && !(node.ChildNodes[4] == null))
                                    {
                                        rowCounter++;
                                        if (rowCounter != 1 && node.ChildNodes[4].InnerText != "n/a" && node.ChildNodes[4].InnerText.Length > 1)
                                        {
                                            StringBuilder sb = new StringBuilder();
                                            sb.Append(node.ChildNodes[4].InnerText);
                                            sb = FixEntities(sb);
                                            sw.WriteLine("<row>" + sb + "</row>");
                                            sb.Clear();
                                        }
                                    }
                                }
                                sw.WriteLine("</extract>");
                            }  
                        }
                        else if (multi == true)
                        {
                            using (var sw = new StreamWriter(exportFile, true))
                            {
                                sw.WriteLine("<?xml version=\"1.0\" ?>");
                                sw.WriteLine("<extract>");
                                foreach (var element in root)
                                {
                                    var node = (XmlElement)element;
                                    if (node.Name == "Row" && !(node.ChildNodes[1] == null) && (!(node.ChildNodes[10] == null) && (node.ChildNodes[10].InnerText != "")) || node.Name == "ss:Row" && !(node.ChildNodes[1] == null) && (!(node.ChildNodes[10] == null) && (node.ChildNodes[10].InnerText != "")))
                                    {
                                        rowCounter++;
                                        if (rowCounter != 1 && node.ChildNodes[1].InnerText != "n/a" && node.ChildNodes[1].InnerText != "Source" && node.ChildNodes[1].InnerText.Length > 1)
                                        {
                                            StringBuilder sb = new StringBuilder();
                                            sb.Append(node.ChildNodes[1].InnerText);
                                            sb = FixEntities(sb);
                                            sw.WriteLine("<row>" + sb + "</row>");
                                            sb.Clear();
                                        }
                                    }
                                }
                                sw.WriteLine("</extract>");
                            }
                        }
                    }
                }
                catch (Exception eeeex)
                {
                    MessageBox.Show("Error! " + eeeex.Message + " Extraction process terminated.");
                    return;
                }
                

            }
        }

        private StringBuilder FixEntities(StringBuilder sb)
        {
            sb.Replace("<", "&lt;");
            sb.Replace(">", "&gt;");
            sb.Replace("&", "&amp;");
            return sb;
        }
    }
}
