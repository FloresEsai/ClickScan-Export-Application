using CMImaging;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.IO;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        private Connection? _conn; // Connection object for database interactions //(FOUND AUTOMATICALLY AND FOUND IN CONFIG FILE)//
        public string? _datasource = "PDS ClickScan"; // The data source for the database connection
        public string? _dname; // The name of the drawer in the database
        public string? _maindest; // The main destination for export files

        // Import the necessary WinAPI functions
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FreeConsole();

        public MainWindow()
        {
            InitializeComponent();

            // Attach console to the WPF application FOR DEBUGGING
            AllocConsole();

            // Attempt to connect to the database if a data source is provided
            if (_datasource != null)
            {
                TryConnectToDatabase();
            }
            Console.WriteLine($"Active Data Source set to: {_datasource}");

            // Populate the datasources and drawers with available options from the current connection
            PopulateDatasourcesComboBox();
            PopulateDrawersComboBox();
        }

        /// <summary>
        /// Closes Console when program is closed out
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // Free the console when the application is closed
            FreeConsole();
        }

        /// <summary>
        ///  Ensure that the combo box is populated with the drawer options before any selection change event occurs
        /// </summary>
        private void PopulateDatasourcesComboBox()
        {
            var sources = _conn.Sources;
            DatasourcesComboBox.Items.Clear(); // Clear existing items

            foreach (var source in sources)
            {
                DatasourcesComboBox.Items.Add(source);
            }
        }

        /// <summary>
        ///  Ensure that the combo box is populated with the drawer options before any selection change event occurs
        /// </summary>
        private void PopulateDrawersComboBox()
        {
            var drawers = _conn.Drawers;
            DrawersComboBox.Items.Clear(); // Clear existing items

            foreach (var drawer in drawers)
            {
                DrawersComboBox.Items.Add(drawer);
            }
        }

        /// <summary>
        /// Attempts to connect to the database using the provided data source.
        /// </summary>
        private void TryConnectToDatabase()
        {
            try
            {
                // Initialize and connect the database connection
                _conn = new Connection();
                _conn.ConnectDB(_datasource);
                System.Windows.MessageBox.Show($"Current Data Source set to: {_datasource}", "Datasource Set");
            }
            catch (Exception)
            {
                System.Windows.MessageBox.Show($"Connection to Data Source was unsuccessful: ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Event handler for the datasources combo box selection change.
        /// Sets the current datasource in the database connection.
        /// </summary>
        private void DatasourcesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get the data sources from the connection
            var sources = _conn.Sources;

            // Populate the combo box with data sources if not already populated
            if (DatasourcesComboBox.Items.Count == 0)
            {
                foreach (var source in sources)
                {
                    DatasourcesComboBox.Items.Add(source);
                }
            }

            // Handle the selection change event
            if (DatasourcesComboBox.SelectedItem is string selectedDatasource)
            {
                try
                {
                    // Initialize and connect the database connection
                    _conn = new Connection();
                    _conn.ConnectDB(selectedDatasource);
                    _datasource = selectedDatasource;
                    System.Windows.MessageBox.Show($"Current Data Source set to: {_datasource}", "Datasource Set");
                    Console.WriteLine($"Active Data Source set to: {_datasource}");
                }
                catch (Exception)
                {
                    // Show error message if connection fails
                    System.Windows.MessageBox.Show($"Connection to Data Source was unsuccessful: ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


        /// <summary>
        /// Event handler for the drawers combo box selection change.
        /// Sets the active drawer in the database connection.
        /// </summary>
        private void DrawersComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get the drawers from the connection
            var drawers = _conn.Drawers;

            // Populate the combo box with drawers if not already populated
            if (DrawersComboBox.Items.Count == 0)
            {
                foreach (var drawer in drawers)
                {
                    DrawersComboBox.Items.Add(drawer);
                }
            }

            // Handle the selection change event
            if (DrawersComboBox.SelectedItem is Drawer selectedDrawer)
            {
                // Set the active drawer in the database connection
                _dname = selectedDrawer.Name;
                _conn.SetActiveDrawer(_dname);
                System.Windows.MessageBox.Show($"Active drawer set to: {_dname}", "Drawer Set");
                Console.WriteLine($"Active drawer set to: {_dname}");
            }
        }


        /// <summary>
        /// Event handler for the "Browse" button click.
        /// Prompts the user to select an export path.
        /// </summary>
        private void EStart_Click(object sender, RoutedEventArgs e)
        {
            // Configure folder browser dialog box
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a folder to export files to";
                dialog.ShowNewFolderButton = true;

                // Show folder browser dialog box
                DialogResult result = dialog.ShowDialog();

                // Process folder browser dialog box results
                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    // Get the selected folder
                    _maindest = dialog.SelectedPath;

                    // Display the current selected path to the user
                    FilePathTextBox.Text = _maindest;
                    Console.WriteLine("Debugging: " + _maindest + " is the current set export destination");
                }
            }
        }

        /// <summary>
        /// Event handler for the "Export" button click
        /// Initiates the exportation from \\_datasource\\_dname to \\_maindest
        /// </summary>
        private void EXButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Extraction Starting");

            DateTime start = DateTime.Now;
            int totalRecords = 0;
            int totalSuccess = 0;
            int totalFailed = 0;

            Console.WriteLine("Attempting to Query Drawer: " + _dname + " from: " + _datasource);

            List<FolderRecord> records = _conn.RetrieveRecords();

            if (records != null && records.Count > 0)
            {
                totalRecords = records.Count;
                Console.WriteLine("Found " + records.Count + " records");

                EnsureDirectoryExists(_maindest + "Images");

                int pageseq = 1;

                using (StreamWriter writer = new(_maindest + "import.txt"))
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
                            LogError("Error Retrieving Images [" + record.FolderID + "]");
                            continue;
                        }

                        ProcessPages(record, writer, ref pageseq, ref totalSuccess, ref totalFailed);

                        // Update the progress bar
                        ReportProgress((pos + 1, totalRecords, totalSuccess, totalFailed));
                    }
                }
                LogExportResults(start, totalRecords, totalSuccess, totalFailed);
            }
            _conn.DBDisconnect();

            System.Windows.MessageBox.Show("Extraction Finished");
        }

        /// <summary>
        /// Updates progress bar in MainWindow.xaml with each export executed
        /// </summary>
        /// <param name="progress"> Encapsulates the int variables as one accessible variable </param>
        /// <param name="current"> Current number of records exported </param>
        /// <param name="total"> Total number of records to be exported </param>
        /// <param name="success"> Total number of successful records exported </param>
        /// <param name="failed"> Total number of failed records exported </param> 
        private void ReportProgress((int current, int total, int success, int failed) progress)
        {
            int percentage = (int)((double)progress.current / progress.total * 100);

            // Ensure the updates are performed on the UI thread
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                ProgressBar.Value = percentage;
                ProgressText.Text = $"{percentage}%";
                SuccessfulCounterText.Text = progress.success.ToString();
                FailedCounterText.Text = progress.failed.ToString();
                TotalDocText.Text = progress.total.ToString();
            });
        }

        /// <summary>
        /// Ensures that a directory exists, creates it if it does not.
        /// </summary>
        /// <param name="path">The path of the directory to check or create.</param>
        private static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        /// <summary>
        /// Validates if the pages in a record are valid.
        /// </summary>
        /// <param name="record">The record to validate.</param>
        /// <returns>True if valid, otherwise false.</returns>
        private static bool IsValidPages(FolderRecord record)
        {
            return record.Pages != null && record.Pages.Count > 0;
        }

        /// <summary>
        /// Logs an error message to the error log.
        /// </summary>
        /// <param name="message">The error message to log.</param>
        /// TODO: COMBINE THIS METHOD WITH THE LOGEXPORTRESULTS METHOD
        private void LogError(string message)
        {
            using (StreamWriter logger = File.AppendText(_maindest + "error.log"))
            {
                logger.WriteLine(message);
                logger.WriteLine("");
            }
        }

        /// <summary>
        /// Processes the pages of a record and writes the output to the specified writer.
        /// </summary>
        /// <param name="record">The record containing the pages.</param>
        /// <param name="writer">The StreamWriter to write the output to.</param>
        /// <param name="pageseq">The page sequence number.</param>
        /// <param name="totalSuccess">The total number of successful exports.</param>
        /// <param name="totalFailed">The total number of failed exports.</param>
        private void ProcessPages(FolderRecord record, StreamWriter writer, ref int pageseq, ref int totalSuccess, ref int totalFailed)
        {
            try
            {
                foreach (CMImaging.Page recpage in record.Pages)
                {
                    string imgfile = _maindest + "Images\\" + pageseq.ToString("0000") + ".tif";

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
                            LogError($"Invalid Image or Format Type {{ {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation} }}");
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
                            LogError($"Invalid Image or Format Type {{ {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation} }}");
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

        /// <summary>
        /// Processes a PDF page and saves the output to the specified image file.
        /// </summary>
        /// <param name="recpage">The page to process.</param>
        /// <param name="imgfile">The output image file path.</param>
        /// <param name="totalSuccess">The total number of successful exports.</param>
        private void ProcessPdfPage(CMImaging.Page recpage, string imgfile, ref int totalSuccess)
        {
            try
            {
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

        /// <summary>
        /// Logs the results of the export process.
        /// </summary>
        /// <param name="start">The start time of the export process.</param>
        /// <param name="totalRecords">The total number of records processed.</param>
        /// <param name="totalSuccess">The total number of successful exports.</param>
        /// <param name="totalFailed">The total number of failed exports.</param>
        private void LogExportResults(DateTime start, int totalRecords, int totalSuccess, int totalFailed)
        {
            Console.WriteLine("NOTICE: Export file will be overwritten if subsequent exports are made in the same directory.");
            try
            {
                // Ensure the file 'export.log' is created in the _maindest location
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
    } // partial class scope
} // namespace scope