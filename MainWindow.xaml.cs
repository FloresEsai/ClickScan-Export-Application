using CMImaging;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.ComponentModel;

namespace WpfApp1
{
    /// <summary>
    /// The main window of the application.
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        // The connection to the ClickScan database
        private Connection? _conn;

        // The name of the data source
        public string? _datasource = "PDS ClickScan";

        // The name of the document
        public string? _dname;

        // The main destination directory
        public string? _maindest = @"";

        // The background worker for exporting records
        private BackgroundWorker _backgroundWorker;

        // The total number of records
        public int totalRecords = 0;

        // The total number of successful exports
        public int totalSuccess = 0;

        // The total number of failed exports
        public int totalFailed = 0;


        /// <summary>
        /// The number of successful exports.
        /// </summary>
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


        /// <summary>
        /// The number of failed exports.
        /// </summary>
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


        /// <summary>
        /// The total number of documents.
        /// </summary>
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


        /// <summary>
        /// Event that is raised when a property value has changed.
        /// </summary>
        public event PropertyChangedEventHandler? PropertyChanged;


        /// <summary>
        /// Raises the PropertyChanged event for the specified property name.
        /// </summary>
        /// <param name="propertyName">The name of the property that has changed.</param>
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }


        /// <summary>
        /// Allocates a new console for the current process.
        /// </summary>
        /// <returns>True if the operation was successful, false otherwise.</returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();


        /// <summary>
        /// Frees the console associated with the current process.
        /// </summary>
        /// <returns>True if the operation was successful, false otherwise.</returns>
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FreeConsole();


        /// <summary>
        /// Initializes the MainWindow.
        /// Allocates a console, connects to the database if a data source is provided,
        /// populates the data sources combo box, and initializes the background worker.
        /// </summary>
        public MainWindow()
        {
            // Initialize the component
            InitializeComponent();

            // Allocate a console
            AllocConsole();

            // Check if a data source is provided
            if (_datasource != null)
            {
                // Try to connect to the database
                TryConnectToDatabase();
            }

            // Write the active data source to the console
            Console.WriteLine($"Active Data Source set to: {_datasource}");

            // Populate the data sources combo box
            PopulateDatasourcesComboBox();

            // Initialize the background worker
            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.WorkerReportsProgress = true;
            _backgroundWorker.DoWork += _backgroundWorker_DoWork;
            _backgroundWorker.ProgressChanged += _backgroundWorker_ProgressChanged;
            _backgroundWorker.RunWorkerCompleted += _backgroundWorker_RunWorkerCompleted;

            // Set the data context of the window to itself
            DataContext = this;
        }


        /// <summary>
        /// Overrides the OnClosed method to perform cleanup actions when the window is closed.
        /// </summary>
        /// <param name="e">The event data.</param>
        protected override void OnClosed(EventArgs e)
        {
            // Call the base class's OnClosed method to ensure proper cleanup
            base.OnClosed(e);

            // Free the console to release any resources associated with it
            FreeConsole();

            // Disconnect from the database to release any resources associated with the connection
            _conn.DBDisconnect();
        }


        /// <summary>
        /// Populates the DatasourcesComboBox with the available data sources from the Connection class.
        /// </summary>
        private void PopulateDatasourcesComboBox()
        {
            // Get the list of available data sources from the Connection class
            var sources = _conn.Sources;

            // Clear the existing items in the DatasourcesComboBox
            DatasourcesComboBox.Items.Clear();

            // Add each data source to the DatasourcesComboBox
            foreach (var source in sources)
            {
                DatasourcesComboBox.Items.Add(source);
            }
        }


        /// <summary>
        /// Populates the DrawersComboBox with the available drawers from the Connection class.
        /// </summary>
        private void PopulateDrawersComboBox()
        {
            // Get the list of drawers from the Connection class
            var drawers = _conn.Drawers;

            // Clear the existing items in the DrawersComboBox
            DrawersComboBox.Items.Clear();

            // Add each drawer to the DrawersComboBox
            foreach (var drawer in drawers)
            {
                DrawersComboBox.Items.Add(drawer);
            }
        }


        /// <summary>
        /// Tries to connect to the database using the specified data source.
        /// Displays a message box indicating the success or failure of the connection attempt.
        /// </summary>
        private void TryConnectToDatabase()
        {
            try
            {
                // Create a new instance of the Connection class
                _conn = new Connection();

                // Connect to the database using the specified data source
                _conn.ConnectDB(_datasource);

                // Display a message box indicating the successful connection
                System.Windows.MessageBox.Show($"Current Data Source set to: {_datasource}", "Datasource Set");
            }
            catch (Exception)
            {
                // Display a message box indicating the unsuccessful connection attempt
                System.Windows.MessageBox.Show($"Connection to Data Source was unsuccessful: ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// Handles the SelectionChanged event of the DatasourcesComboBox.
        /// Updates the active data source and displays a message box.
        /// If there are no items in the DatasourcesComboBox, it populates it with the available sources.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void DatasourcesComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get the list of available sources
            var sources = _conn.Sources;

            // If the DatasourcesComboBox is empty, populate it with the available sources
            if (DatasourcesComboBox.Items.Count == 0)
            {
                foreach (var source in sources)
                {
                    DatasourcesComboBox.Items.Add(source);
                }
            }

            // Check if a data source is selected
            if (DatasourcesComboBox.SelectedItem is DataSource selectedDatasource)
            {
                // Check if a data source is already active
                if (_datasource != null)
                {
                    try
                    {
                        // Connect to the selected data source
                        _conn.ConnectDB(selectedDatasource.Name);
                        _datasource = selectedDatasource.Name;
                        System.Windows.MessageBox.Show($"Current Data Source set to: {_datasource}", "Datasource Set");
                        Console.WriteLine($"Active Data Source set to: {_datasource}");
                        PopulateDrawersComboBox();
                    }
                    catch (Exception)
                    {
                        // Display an error message if the connection to the data source fails
                        System.Windows.MessageBox.Show($"Connection to Data Source was unsuccessful: ", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    // Display an error message if no data source is selected
                    System.Windows.MessageBox.Show($"No data source selected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


        /// <summary>
        /// Handles the SelectionChanged event of the DrawersComboBox.
        /// Updates the active drawer and displays a message box.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void DrawersComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Get the list of drawers
            var drawers = _conn.Drawers;

            // If the DrawersComboBox is empty, populate it with the drawers
            if (DrawersComboBox.Items.Count == 0)
            {
                foreach (var drawer in drawers)
                {
                    DrawersComboBox.Items.Add(drawer);
                }
            }

            // If an item is selected in the DrawersComboBox
            if (DrawersComboBox.SelectedItem is Drawer selectedDrawer)
            {
                // Set the active drawer name
                _dname = selectedDrawer.Name;

                // Set the active drawer in the connection
                _conn.SetActiveDrawer(_dname);

                // Display a message box with the active drawer name
                System.Windows.MessageBox.Show($"Active drawer set to: {_dname}", "Drawer Set");

                // Log the active drawer name to the console
                Console.WriteLine($"Active drawer set to: {_dname}");
            }
        }


        /// <summary>
        /// Handles the click event of the EStartButton.
        /// Opens a folder browser dialog to select the export destination.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void EStart_Click(object sender, RoutedEventArgs e)
        {
            // Open a folder browser dialog to select the export destination
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                // Set the description and show a new folder button
                dialog.Description = "Set Export Destination";
                dialog.ShowNewFolderButton = true;

                // Show the dialog and get the result
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                // If the dialog was successfully closed with the OK button
                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    // Set the export destination
                    _maindest = dialog.SelectedPath;

                    // Ensure the export destination ends with a directory separator
                    if (!_maindest.EndsWith(Path.DirectorySeparatorChar.ToString()))
                    {
                        _maindest += Path.DirectorySeparatorChar;
                    }

                    // Update the text of the FilePathTextBox
                    FilePathTextBox.Text = _maindest;

                    // Log the selected export destination
                    Console.WriteLine("Debugging: " + _maindest + " is the current set export destination");
                }
            }
        }


        /// <summary>
        /// Handles the click event of the EXButton.
        /// Starts the extraction process if it's not already running.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void EXButton_Click(object sender, RoutedEventArgs e)
        {
            // Show a message box indicating that the extraction has started
            System.Windows.MessageBox.Show("Extraction Starting");

            // Check if the background worker is not already busy
            if (!_backgroundWorker.IsBusy)
            {
                // Start the extraction process by running the background worker
                _backgroundWorker.RunWorkerAsync();
            }
        }


        /// <summary>
        /// This method is called when the background worker starts to process.
        /// It retrieves records from the database, processes each record, and logs the results.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event arguments.</param>
        private void _backgroundWorker_DoWork(object? sender, DoWorkEventArgs e)
        {
            // Get the start time
            DateTime start = DateTime.Now;

            // Log the attempt to query the database
            Console.WriteLine("Attempting to Query Drawer: " + _dname + " from: " + _datasource);

            // Retrieve records from the database
            List<FolderRecord> records = _conn.RetrieveRecords();

            // If there are records, process them
            if (records != null && records.Count > 0)
            {
                // Set the total number of records
                totalRecords = records.Count;

                // Log the number of records found
                Console.WriteLine("Found " + records.Count + " records");

                // Ensure the "Images" directory exists
                EnsureDirectoryExists(Path.Combine(_maindest, "Images"));

                // Initialize the page sequence
                int pageseq = 1;

                // Open a file to write the import data
                using (StreamWriter writer = new(Path.Combine(_maindest, "import.txt")))
                {
                    // Process each record
                    for (int pos = 0; pos < records.Count; pos++)
                    {
                        // Get the current record
                        FolderRecord record = records[pos];

                        // Log the progress
                        Console.WriteLine("Extracting record " + (pos + 1) + " of " + records.Count);

                        // Get the pages for the record
                        record.Pages = _conn.GetPages(record.FolderID);

                        // If the pages are not valid, log an error and skip the record
                        if (!IsValidPages(record))
                        {
                            totalFailed++;
                            totalRecords--;
                            LogError("Error Retrieving Images [" + record.FolderID + "]", Path.Combine(_maindest, "ErrorLog.txt"));
                            continue;
                        }

                        // Process the pages
                        ProcessPages(record, writer, ref pageseq, ref totalSuccess, ref totalFailed);

                        // Update the properties
                        SuccessfulExports = totalSuccess;
                        FailedExports = totalFailed;
                        TotalDocuments = totalRecords;

                        // Calculate the progress percentage
                        int progressPercentage = (int)((pos + 1) / (double)totalRecords * 100);

                        // Report the progress to the background worker
                        _backgroundWorker.ReportProgress(progressPercentage);
                    }
                }

                // Log the export results
                LogExportResults(start, totalRecords, totalSuccess, totalFailed);
            }
        }


        /// <summary>
        /// This method is called whenever the progress of a background worker changes.
        /// It updates the value of the ProgressBar control to match the progress percentage.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event arguments containing the progress percentage.</param>
        private void _backgroundWorker_ProgressChanged(object? sender, ProgressChangedEventArgs e)
        {
            // Update the value of the ProgressBar control to match the progress percentage
            ProgressBar.Value = e.ProgressPercentage;
        }


        /// <summary>
        /// Handles the completion of the background worker.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The event data.</param>
        private void _backgroundWorker_RunWorkerCompleted(object? sender, RunWorkerCompletedEventArgs e)
        {
            // Show a message box indicating that the extraction has finished
            System.Windows.MessageBox.Show("Extraction Finished");
        }


        /// <summary>
        /// Ensures that a directory exists at the specified path.
        /// </summary>
        /// <param name="path">The path of the directory.</param>
        /// <returns>True if the directory exists or was successfully created, false otherwise.</returns>
        private static bool EnsureDirectoryExists(string path)
        {
            // Check if the directory already exists
            if (!Directory.Exists(path))
            {
                try
                {
                    // Attempt to create the directory
                    Directory.CreateDirectory(path);
                }
                catch (Exception e)
                {
                    // Log the error and return false if an exception occurs
                    Console.WriteLine($"Failed to create directory: {e.Message}");
                    return false;
                }
            }
            // Return true if the directory exists or was successfully created
            return Directory.Exists(path);
        }


        /// <summary>
        /// Checks if the given FolderRecord has valid pages.
        /// </summary>
        /// <param name="record">The FolderRecord to check.</param>
        /// <returns>True if the record has valid pages, false otherwise.</returns>
        private static bool IsValidPages(FolderRecord record)
        {
            // Check if the Pages property of the record is not null and has at least one page.
            return record.Pages != null && record.Pages.Count > 0;
        }


        /// <summary>
        /// Logs an error message to a log file.
        /// </summary>
        /// <param name="message">The error message to log.</param>
        /// <param name="logFilePath">The path to the log file.</param>
        private void LogError(string message, string logFilePath)
        {
            // Check if the log file exists, and create it if it doesn't
            if (!File.Exists(logFilePath))
            {
                File.Create(logFilePath).Dispose();
            }

            // Open the log file for writing
            using (StreamWriter logger = File.AppendText(logFilePath))
            {
                // Write the error message to the log file
                logger.WriteLine(message);

                // Write an empty line to separate the error message from the next entry
                logger.WriteLine("");
            }
        }


        /// <summary>
        /// Processes the pages of a given FolderRecord and writes the image file paths to a StreamWriter.
        /// </summary>
        /// <param name="record">The FolderRecord containing the pages to be processed.</param>
        /// <param name="writer">The StreamWriter used to write the image file paths.</param>
        /// <param name="pageseq">A reference to an integer representing the current page sequence number.</param>
        /// <param name="totalSuccess">A reference to an integer representing the total number of successful exports.</param>
        /// <param name="totalFailed">A reference to an integer representing the total number of failed exports.</param>
        private void ProcessPages(FolderRecord record, StreamWriter writer, ref int pageseq, ref int totalSuccess, ref int totalFailed)
        {
            try
            {
                foreach (CMImaging.Page recpage in record.Pages)
                {
                    // Generate the image file path
                    string imgfile = Path.Combine(_maindest, "Images", pageseq.ToString("0000") + ".tif");

                    // Check if the file already exists
                    while (File.Exists(imgfile))
                    {
                        pageseq++;
                        imgfile = Path.Combine(_maindest, "Images", pageseq.ToString("0000") + ".tif");
                    }

                    if (recpage.FileType == FileType.ImageFile)
                    {
                        if (recpage.FileType == FileType.PDF)
                        {
                            // Process PDF page
                            ProcessPdfPage(recpage, imgfile, ref totalSuccess);
                        }
                        else if (recpage.ImageType == ImageFileType.TIFF || recpage.ImageType == ImageFileType.JPEG || recpage.ImageType == ImageFileType.BMP || recpage.ImageType == ImageFileType.PNG)
                        {
                            // Save image file
                            System.Drawing.Image tmpimg = System.Drawing.Image.FromFile(recpage.FileLocation);
                            tmpimg.Save(imgfile, _conn.EncoderInfo, _conn.ImgEncParams);
                            totalSuccess++;
                        }
                        else
                        {
                            // Log error for invalid image or format type
                            Console.WriteLine($"A File has failed to be exported. Description(ID:ImageType:FileType:Location): {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation}");
                            totalFailed++;
                            LogError($"Invalid Image or Format Type {{ {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation} }}", Path.Combine(_maindest, "ErrorLog.txt"));
                        }
                    }
                    else
                    {
                        if (recpage.FileType == FileType.PDF)
                        {
                            // Process PDF page
                            ProcessPdfPage(recpage, imgfile, ref totalSuccess);
                        }
                        else if (recpage.FileType == FileType.ImageFile)
                        {
                            // Save the image file
                            System.Drawing.Image tmpimg = System.Drawing.Image.FromFile(recpage.FileLocation);
                            tmpimg.Save(imgfile, _conn.EncoderInfo, _conn.ImgEncParams);
                            totalSuccess++;
                        }
                        else
                        {
                            // Log error for invalid image or format type
                            Console.WriteLine($"Failed Export. Not of type 'ImageFile': {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation}");
                            LogError($"Invalid Image or Format Type {{ {record.FolderID} : {recpage.ImageType} : {recpage.FileType} : {recpage.FileLocation} }}", Path.Combine(_maindest, "ErrorLog.txt"));
                            totalFailed++;
                        }
                    }

                    // Write the image file path to the StreamWriter
                    writer.WriteLine((recpage.Number == 1 ? record.DelimitedIndex : "") + "@" + imgfile);
                    pageseq++;
                }
            }
            catch (Exception ex)
            {
                // Display error message if an exception occurs
                System.Windows.MessageBox.Show($"Error in File: MainWindow.xaml.cs : Method: ProcessPages : {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// Processes a PDF page by rendering it and saving it as an image file.
        /// </summary>
        /// <param name="recpage">The PDF page to process.</param>
        /// <param name="imgfile">The path and filename for the output image file.</param>
        /// <param name="totalSuccess">The total number of successful image exports.</param>
        private void ProcessPdfPage(CMImaging.Page recpage, string imgfile, ref int totalSuccess)
        {
            try
            {
                // Generate a unique filename for the image file
                while (File.Exists(imgfile))
                {
                    int index = imgfile.LastIndexOf('.');
                    string baseName = imgfile.Substring(0, index);
                    string extension = imgfile.Substring(index);
                    int number = int.Parse(baseName.Substring(baseName.Length - 4));
                    number++;
                    imgfile = baseName + number.ToString("0000") + extension;
                }

                // Load the PDF document
                using (PdfiumViewer.PdfDocument pdfdoc = PdfiumViewer.PdfDocument.Load(recpage.FileLocation))
                {
                    // Render each page of the PDF document as an image
                    for (int p = 0; p < pdfdoc.PageCount; p++)
                    {
                        using (var tmpImg = pdfdoc.Render(p, (int)(pdfdoc.PageSizes[p].Width / 0.24f), (int)(pdfdoc.PageSizes[p].Height / 0.24f), 600, 600, false))
                        {
                            try
                            {
                                // Save the image file
                                tmpImg.Save(imgfile, _conn.EncoderInfo, _conn.ImgEncParams);
                                totalSuccess++;
                            }
                            catch (Exception)
                            {
                                totalFailed++;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Show a message box indicating that an error occurred while processing the PDF page
                System.Windows.MessageBox.Show($"Error in File: MainWindow.xaml.cs : Method: ProcessPdfPage", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>
        /// Logs the results of an export operation to a file.
        /// </summary>
        /// <param name="start">The start time of the export operation.</param>
        /// <param name="totalRecords">The total number of records exported.</param>
        /// <param name="totalSuccess">The number of successful exports.</param>
        /// <param name="totalFailed">The number of failed exports.</param>
        private void LogExportResults(DateTime start, int totalRecords, int totalSuccess, int totalFailed)
        {
            // Print a notice about file overwrite
            Console.WriteLine("NOTICE: Export file will be overwritten if subsequent exports are made in the same directory.");

            try
            {
                // Define the path for the log file
                string logFilePath = Path.Combine(_maindest, "export.log");

                // Create the log file if it doesn't exist
                if (!File.Exists(logFilePath))
                {
                    using (File.Create(logFilePath)) { }
                }

                // Open the log file for writing
                using (StreamWriter logger = File.AppendText(logFilePath))
                {
                    // Write the header for the log file
                    logger.WriteLine("================================================================");
                    logger.WriteLine("                           Extraction Information");
                    logger.WriteLine("================================================================");

                    // Write the start time of the export
                    logger.WriteLine("   Extraction Started: " + start.ToString("MM/dd/yyyy hh:mm:ss tt"));

                    // Write the drawer information
                    logger.WriteLine("   Drawer Id: " + _conn.ActiveDrawer.ID);
                    logger.WriteLine("   Drawer Name: " + _conn.ActiveDrawer.Name);

                    // Write the export destination
                    logger.WriteLine("   Export Destination: " + _maindest);

                    // Write the total number of records exported
                    logger.WriteLine("   Total Records: " + totalRecords);

                    // Write the number of successful exports
                    logger.WriteLine("   Total Successful Exports: " + totalSuccess);

                    // Write the number of failed exports if any
                    if (totalFailed > 0)
                        logger.WriteLine("   Errors: " + totalFailed);

                    // Write the end time of the export
                    logger.WriteLine("   Extraction Completed: " + DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"));

                    // Write the footer for the log file
                    logger.WriteLine("================================================================");
                    logger.WriteLine("================================================================");

                    // Flush the writer to ensure all data is written to the file
                    logger.Flush();
                }

                // Show a message box indicating that the export log has been created
                System.Windows.MessageBox.Show("Export log created");
            }
            catch (Exception)
            {
                // Show a message box indicating that an error occurred while logging the export results
                System.Windows.MessageBox.Show($"Error in File: MainWindow.xaml.cs : Method: LogExportResults", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
