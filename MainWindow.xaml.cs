using CMImaging;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.ComponentModel;

namespace WpfApp1
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private Connection? _conn;
        public string? _datasource = "PDS ClickScan";
        public string? _dname;
        public string? _maindest = @"";
        private BackgroundWorker _backgroundWorker;
        public int totalRecords = 0;
        public int totalSuccess = 0;
        public int totalFailed = 0;

        private int _successfulExports;
        public int SuccessfulExports
        {
            get => _successfulExports;
            set
            {
                _successfulExports = totalSuccess;
                OnPropertyChanged(nameof(SuccessfulExports));
            }
        }

        private int _failedExports;
        public int FailedExports
        {
            get => _failedExports;
            set
            {
                _failedExports = totalFailed;
                OnPropertyChanged(nameof(FailedExports));
            }
        }

        private int _totalDocuments;
        public int TotalDocuments
        {
            get => _totalDocuments;
            set
            {
                _totalDocuments = totalRecords;
                OnPropertyChanged(nameof(TotalDocuments));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FreeConsole();

        public MainWindow()
        {
            InitializeComponent();
            AllocConsole();

            if (_datasource != null)
            {
                TryConnectToDatabase();
            }
            Console.WriteLine($"Active Data Source set to: {_datasource}");

            PopulateDatasourcesComboBox();

            // Initialize BackgroundWorker
            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.WorkerReportsProgress = true;
            _backgroundWorker.DoWork += _backgroundWorker_DoWork;
            _backgroundWorker.ProgressChanged += _backgroundWorker_ProgressChanged;
            _backgroundWorker.RunWorkerCompleted += _backgroundWorker_RunWorkerCompleted;

            DataContext = this;
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            FreeConsole();
            _conn.DBDisconnect();
        }

        private void PopulateDatasourcesComboBox()
        {
            var sources = _conn.Sources;
            DatasourcesComboBox.Items.Clear();
            foreach (var source in sources)
            {
                DatasourcesComboBox.Items.Add(source);
            }
        }

        private void PopulateDrawersComboBox()
        {
            var drawers = _conn.Drawers;
            DrawersComboBox.Items.Clear();
            foreach (var drawer in drawers)
            {
                DrawersComboBox.Items.Add(drawer);
            }
        }

        private void TryConnectToDatabase()
        {
            try
            {
                _conn = new Connection();
                _conn.ConnectDB(_datasource);
                System.Windows.MessageBox.Show($"Current Data Source set to: {_datasource}", "Datasource Set");
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show($"Connection to Data Source was unsuccessful: ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DatasourcesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var sources = _conn.Sources;
            if (DatasourcesComboBox.Items.Count == 0)
            {
                foreach (var source in sources)
                {
                    DatasourcesComboBox.Items.Add(source);
                }
            }

            if (DatasourcesComboBox.SelectedItem is DataSource selectedDatasource)
            {
                if (_datasource != null)
                {
                    try
                    {
                        _conn.ConnectDB(selectedDatasource.Name);
                        _datasource = selectedDatasource.Name;
                        System.Windows.MessageBox.Show($"Current Data Source set to: {_datasource}", "Datasource Set");
                        Console.WriteLine($"Active Data Source set to: {_datasource}");
                        PopulateDrawersComboBox();
                    }
                    catch (Exception)
                    {
                        System.Windows.MessageBox.Show($"Connection to Data Source was unsuccessful: ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    System.Windows.MessageBox.Show($"No data source selected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void DrawersComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var drawers = _conn.Drawers;
            if (DrawersComboBox.Items.Count == 0)
            {
                foreach (var drawer in drawers)
                {
                    DrawersComboBox.Items.Add(drawer);
                }
            }

            if (DrawersComboBox.SelectedItem is Drawer selectedDrawer)
            {
                _dname = selectedDrawer.Name;
                _conn.SetActiveDrawer(_dname);
                System.Windows.MessageBox.Show($"Active drawer set to: {_dname}", "Drawer Set");
                Console.WriteLine($"Active drawer set to: {_dname}");
            }
        }

        private void EStart_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Set Export Destination";
                dialog.ShowNewFolderButton = true;
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    _maindest = dialog.SelectedPath;

                    if (!_maindest.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    {
                        _maindest += Path.DirectorySeparatorChar;
                    }

                    FilePathTextBox.Text = _maindest;
                    Console.WriteLine("Debugging: " + _maindest + " is the current set export destination");
                }
            }
        }

        private void EXButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Extraction Starting");

            if (!_backgroundWorker.IsBusy)
            {
                _backgroundWorker.RunWorkerAsync();
            }
        }

        private void _backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            DateTime start = DateTime.Now;


            Console.WriteLine("Attempting to Query Drawer: " + _dname + " from: " + _datasource);

            List<FolderRecord> records = _conn.RetrieveRecords();

            if (records != null && records.Count > 0)
            {
                totalRecords = records.Count;
                Console.WriteLine("Found " + records.Count + " records");

                EnsureDirectoryExists(Path.Combine(_maindest, "Images"));

                int pageseq = 1;

                using (StreamWriter writer = new(Path.Combine(_maindest, "import.txt")))
                {
                    for (int pos = 0; pos < records.Count; pos++)
                    {
                        FolderRecord record = records[pos];
                        Console.WriteLine("Extracting record " + (pos + 1) + " of " + records.Count);

                        record.Pages = _conn.GetPages(record.FolderID);

                        if (!IsValidPages(record))
                        {
                            totalFailed++;
                            totalRecords--;
                            LogError("Error Retrieving Images [" + record.FolderID + "]", Path.Combine(_maindest, "ErrorLog.txt"));
                            continue;
                        }

                        ProcessPages(record, writer, ref pageseq, ref totalSuccess, ref totalFailed);

                        // Update the properties
                        SuccessfulExports = totalSuccess;
                        FailedExports = totalFailed;
                        TotalDocuments = totalRecords;

                        int progressPercentage = (int)((pos + 1) / (double)totalRecords * 100);
                        _backgroundWorker.ReportProgress(progressPercentage);
                    }
                }
                LogExportResults(start, totalRecords, totalSuccess, totalFailed);
            }
        }

        private void _backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            ProgressBar.Value = e.ProgressPercentage;
        }

        private void _backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            System.Windows.MessageBox.Show("Extraction Finished");
        }

        private static bool EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to create directory: {e.Message}");
                    return false;
                }
            }
            return Directory.Exists(path);
        }

        private static bool IsValidPages(FolderRecord record)
        {
            return record.Pages != null && record.Pages.Count > 0;
        }

        private void LogError(string message, string logFilePath)
        {
            using (StreamWriter logger = File.AppendText(logFilePath))
            {
                logger.WriteLine(message);
                logger.WriteLine("");
            }
        }

        private void ProcessPages(FolderRecord record, StreamWriter writer, ref int pageseq, ref int totalSuccess, ref int totalFailed)
        {
            try
            {
                foreach (CMImaging.Page recpage in record.Pages)
                {
                    string imgfile = Path.Combine(_maindest, "Images", pageseq.ToString("0000") + ".tif");

                    while (File.Exists(imgfile))
                    {
                        pageseq++;
                        imgfile = Path.Combine(_maindest, "Images", pageseq.ToString("0000") + ".tif");
                    }

                    if (recpage.FileType == FileType.ImageFile)
                    {
                        if (recpage.FileType == FileType.PDF)
                        {
                            ProcessPdfPage(recpage, imgfile, ref totalSuccess);
                        }
                        else if (recpage.ImageType == ImageFileType.TIFF || recpage.ImageType == ImageFileType.JPEG || recpage.ImageType == ImageFileType.BMP || recpage.ImageType == ImageFileType.PNG)
                        {
                            System.Drawing.Image tmpimg = System.Drawing.Image.FromFile(recpage.FileLocation);
                            tmpimg.Save(imgfile, _conn.EncoderInfo, _conn.ImgEncParams);
                            totalSuccess++;
                        }
                        else
                        {
                            Console.WriteLine($"A File has failed to be exported. Description(ID:ImageType:FileType:Location): {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation}");
                            totalFailed++;
                            LogError($"Invalid Image or Format Type {{ {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation} }}", Path.Combine(_maindest, "ErrorLog.txt"));
                        }
                    }
                    else
                    {
                        if (recpage.FileType == FileType.PDF)
                        {
                            ProcessPdfPage(recpage, imgfile, ref totalSuccess);
                        }
                        else if (recpage.FileType == FileType.ImageFile)
                        {
                            System.Drawing.Image tmpimg = System.Drawing.Image.FromFile(recpage.FileLocation);
                            tmpimg.Save(imgfile, _conn.EncoderInfo, _conn.ImgEncParams);
                            totalSuccess++;
                        }
                        else
                        {
                            Console.WriteLine($"Failed Export. Not of type 'ImageFile': {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation}");
                            LogError($"Invalid Image or Format Type {{ {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation} }}", Path.Combine(_maindest, "ErrorLog.txt"));
                        }
                    }

                    writer.WriteLine((recpage.Number == 1 ? record.DelimitedIndex : "") + "@" + imgfile);
                    pageseq++;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error in File: MainWindow.xaml.cs : Method: ProcessPages : {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ProcessPdfPage(CMImaging.Page recpage, string imgfile, ref int totalSuccess)
        {
            try
            {
                while (File.Exists(imgfile))
                {
                    int index = imgfile.LastIndexOf('.');
                    string baseName = imgfile.Substring(0, index);
                    string extension = imgfile.Substring(index);
                    int number = int.Parse(baseName.Substring(baseName.Length - 4));
                    number++;
                    imgfile = baseName + number.ToString("0000") + extension;
                }

                using (PdfiumViewer.PdfDocument pdfdoc = PdfiumViewer.PdfDocument.Load(recpage.FileLocation))
                {
                    for (int p = 0; p < pdfdoc.PageCount; p++)
                    {
                        using (var tmpImg = pdfdoc.Render(p, (int)(pdfdoc.PageSizes[p].Width / 0.24f), (int)(pdfdoc.PageSizes[p].Height / 0.24f), 600, 600, false))
                        {
                            tmpImg.Save(imgfile, _conn.EncoderInfo, _conn.ImgEncParams);
                            totalSuccess++;
                        }
                    }
                }
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show($"Error in File: MainWindow.xaml.cs : Method: ProcessPdfPage", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LogExportResults(DateTime start, int totalRecords, int totalSuccess, int totalFailed)
        {
            Console.WriteLine("NOTICE: Export file will be overwritten if subsequent exports are made in the same directory.");
            try
            {
                string logFilePath = Path.Combine(_maindest, "export.log");
                if (!File.Exists(logFilePath))
                {
                    using (File.Create(logFilePath)) { }
                }

                using (StreamWriter logger = File.AppendText(logFilePath))
                {
                    logger.WriteLine("================================================================");
                    logger.WriteLine("                                                     Extraction Information");
                    logger.WriteLine("================================================================");
                    logger.WriteLine("     Extraction Started: " + start.ToString("MM/dd/yyyy hh:mm:ss tt"));
                    logger.WriteLine("");
                    logger.WriteLine("     Drawer Id: " + _conn.ActiveDrawer.ID);
                    logger.WriteLine("     Drawer Name: " + _conn.ActiveDrawer.Name);
                    logger.WriteLine("");
                    logger.WriteLine("     Export Destination: " + _maindest);
                    logger.WriteLine("");
                    logger.WriteLine("     Total Records: " + totalRecords);
                    logger.WriteLine("     Total Successful Exports: " + totalSuccess);
                    if (totalFailed > 0)
                        logger.WriteLine("     Errors: " + totalFailed);
                    logger.WriteLine("");
                    logger.WriteLine("     Extraction Completed: " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"));
                    logger.WriteLine("================================================================");
                    logger.WriteLine("================================================================");
                }
                System.Windows.MessageBox.Show("Export log created");
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show($"Error in File: MainWindow.xaml.cs : Method: LogExportResults", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
