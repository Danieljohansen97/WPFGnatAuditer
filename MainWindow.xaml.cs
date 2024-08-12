using ExcelDataReader;
using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System.Data;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace WPFGnatAuditer
{
    public partial class MainWindow : Window
    {
        // Configuration
        private static readonly string folderPath = "";
        private static readonly string fileName = "";
        private string fullPath = Path.Combine(folderPath, fileName);
        private string connectionString = "";
        private List<CiEntry> ciEntries = new();

        public MainWindow()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void ExtractExcelButton_Click(object sender, RoutedEventArgs e)
        {
            StatusEllipse1.Fill = GetSolidColor("Orange");
            Log("Starting data extraction...");

            try
            {
                ExtractExcelData();
            }
            catch (Exception ex)
            {
                StatusEllipse1.Fill = GetSolidColor("Red");
                LogError("ERROR OCCURRED WHILE EXTRACTING DATA:", ex);
            }
        }

        private static SolidColorBrush GetSolidColor(string color)
        {
            switch (color)
            {
                case "Orange":
                    return new SolidColorBrush(Colors.Orange);
                case "Red":
                    return new SolidColorBrush(Colors.Red);
                case "Green":
                    return new SolidColorBrush(Colors.Green);
                default:
                    return new SolidColorBrush(Colors.Black);
            }
        }

        private void UpdateDatabaseButton_Click(object sender, RoutedEventArgs e)
        {
            StatusEllipse2.Fill = GetSolidColor("Orange");

            try
            {
                UpdateDatabase();
            }
            catch (Exception ex)
            {
                StatusEllipse2.Fill = GetSolidColor("Red");
                LogError("ERROR ENCOUNTERED:", ex);
            }
        }

        // Initialize database connection string
        private void InitializeConnectionString()
        {
            connectionString = $"Server={ServerAddressTextBox.Text};Database={DatabaseNameTextBox.Text};Uid={UsernameTextBox.Text};Pwd={PasswordBox.Password};";
        }

        // Update database function
        private void UpdateDatabase()
        {
            Log("DB Update started...");
            InitializeConnectionString();

            using var connection = new MySqlConnection(connectionString);
            connection.Open();

            using var transaction = connection.BeginTransaction();
            using var command = connection.CreateCommand();
            command.Transaction = transaction;

            int totalAffectedRows = 0;

            try
            {
                command.CommandText = @"
                    UPDATE CI_ENTRIES 
                    SET 
                        LOCATION = @Location, 
                        SPECIFIC_LOCATION = @SpecificLocation, 
                        SUBZONE = @SubZone, 
                        SITE = @Site, 
                        SUBSITE = @SubSite, 
                        COMPONENT = @Component, 
                        SUBCOMPONENT = @SubComponent, 
                        CI_PRIORITY = @CiPriority, 
                        TYPE_ID = @Type, 
                        STATE_ID = @State, 
                        CI_DESC = @CiDescription, 
                        CI_NAME = @CiName
                    WHERE CI_ENTRIES_ID = @CiEntriesId;";

                foreach (var ci in ciEntries)
                {
                    AddParameters(command, ci);

                    int affectedRows = command.ExecuteNonQuery();
                    totalAffectedRows += affectedRows;
                }

                transaction.Commit();
                Log($"Update completed. Total affected rows: {totalAffectedRows}.");
                StatusEllipse2.Fill = GetSolidColor("Green");
                UpdateDatabaseButton.IsEnabled = false;
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                LogError("DB Update failed, transaction rolled back.", ex);
            }
        }

        private void AddParameters(MySqlCommand command, CiEntry ci)
        {
            command.Parameters.Clear();
            command.Parameters.AddWithValue("@CiEntriesId", ci.CiEntriesId);
            command.Parameters.AddWithValue("@Location", ci.Location);
            command.Parameters.AddWithValue("@SpecificLocation", ci.SpecificLocation);
            command.Parameters.AddWithValue("@SubZone", ci.SubZone);
            command.Parameters.AddWithValue("@Site", ci.Site);
            command.Parameters.AddWithValue("@SubSite", ci.SubSite);
            command.Parameters.AddWithValue("@Component", ci.Component);
            command.Parameters.AddWithValue("@SubComponent", ci.SubComponent);
            //command.Parameters.AddWithValue("@Node", ci.Node);
            //command.Parameters.AddWithValue("@ProbeSc", ci.ProbeSc);
            command.Parameters.AddWithValue("@CiPriority", ci.CiPriority);
            command.Parameters.AddWithValue("@Type", ci.Type);
            command.Parameters.AddWithValue("@State", ci.State);
            command.Parameters.AddWithValue("@CiDescription", ci.CiDescription);
            command.Parameters.AddWithValue("@CiName", ci.CiName);
        }

        // Extract Excel data function
        private void ExtractExcelData()
        {
            using var stream = File.Open(fullPath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var table = result.Tables[0]; // Assuming the data is in the first sheet
            var firstRow = table.Rows[0];

            int goodItems = 0, badItems = 0;

            if (CheckFirstRowValid(firstRow))
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    var row = table.Rows[i];
                    if (IsRowValid(row))
                    {
                        goodItems++;
                        try
                        {
                            ciEntries.Add(CreateCiEntryFromRow(row));
                        }
                        catch (Exception ex)
                        {
                            LogError($"Failed to create CI entry from row {i}", ex);
                        }
                    }
                    else
                    {
                        badItems++;
                    }
                }
            }
            else
            {
                Log("Check specifications.");
            }

            Log($"Extraction stats\nRows with CI_ENTRY_ID: {goodItems}\nMissing CI_ENTRY_ID: {badItems}");
            Log("Extraction complete, ready to insert into DB.");
            UpdateDatabaseButton.IsEnabled = true;
            ExtractExcelButton.IsEnabled = false;
            StatusEllipse1.Fill = GetSolidColor("Green");
        }

        private bool CheckFirstRowValid(DataRow row)
        {
            if 
                (
                    row[0].ToString() == "R Location"
                    && row[1].ToString() == "R Specific Location"
                    && row[2].ToString() == "R SubZone"
                    && row[3].ToString() == "R Site"
                    && row[4].ToString() == "R SubSite"
                    && row[5].ToString() == "R Component"
                    && row[6].ToString() == "R SubComponent"
                    && row[7].ToString() == "R Node"
                    && row[8].ToString() == "R Probe SC"
                    && row[9].ToString() == "R CI Priority"
                    && row[10].ToString() == "R Type"
                    && row[11].ToString() == "R State"
                    && row[12].ToString() == "R CI Description"
                    && row[13].ToString() == "R CI Name"
                    && row[14].ToString() == "GN CI_ENTRIES_ID"
                )
            {
                Log("Rows are in order, data extraction begins.");
                return true;
            }
            else
            {
                Log("Rows are not ordered correctly, extraction stopped.");
                return false;
            }
        }
        private bool IsRowValid(DataRow row)
        {
            return !string.IsNullOrWhiteSpace(row[14]?.ToString()) && row[14].ToString().All(char.IsDigit);
        }

        private CiEntry CreateCiEntryFromRow(DataRow row)
        {
            return new CiEntry(
                int.Parse(row[14]?.ToString()),
                row[0]?.ToString(),
                row[1]?.ToString(),
                row[2]?.ToString(),
                row[3]?.ToString(),
                row[4]?.ToString(),
                row[5]?.ToString(),
                row[6]?.ToString(),
                row[7]?.ToString(),
                row[8]?.ToString(),
                MapCiPriority(row[9]?.ToString()),
                MapCiType(row[10]?.ToString()),
                MapCiState(row[11]?.ToString()),
                row[12]?.ToString(),
                row[13]?.ToString()
            );
        }

        static int MapCiPriority(string priority) => priority switch
        {
            "Critical" => 4,
            "High" => 3,
            "Medium" => 2,
            "Low" => 1,
            _ => 1
        };

        static int MapCiType(string type) => type switch
        {
            "Enrichment Testing" => 1,
            "AMB" => 2,
            "ATR" => 3,
            "CAM" => 4,
            "CCGW" => 5,
            "CEB" => 6,
            "Channel" => 7,
            "Controller" => 8,
            "Conventional" => 9,
            "Core Router" => 10,
            "Database Server" => 11,
            "DIU" => 12,
            "Environmental" => 13,
            "Exit Router" => 14,
            "GAS Server" => 15,
            "Gateway Router" => 16,
            "LAN Switch" => 17,
            "Link" => 18,
            "Logging Recorder" => 19,
            "MGEG" => 20,
            "Microwave" => 21,
            "Moscad Server" => 22,
            "NTP" => 23,
            "OP" => 24,
            "QUANTAR" => 25,
            "RDM" => 26,
            "Router" => 27,
            "RTU" => 28,
            "Site" => 29,
            "Statistical Server" => 30,
            "Switch" => 31,
            "TENSR" => 32,
            "Terminal Server" => 33,
            "TRAK" => 34,
            "UCS" => 35,
            "UEM" => 36,
            "VMS" => 37,
            "VPM" => 38,
            "ZDS" => 39,
            "Zone Controller" => 40,
            "Gateway Unit" => 41,
            "Data Basestation" => 42,
            "Agent" => 43,
            "Camera" => 44,
            "Infrastructure(CHI CAM)" => 45,
            "LTE" => 46,
            "Network Device" => 47,
            "Logging Replay Station" => 48,
            "Network Address" => 49,
            "Generic Node" => 50,
            "Call Processor" => 51,
            "Data Processing" => 52,
            "Domain Controller" => 53,
            "Backup Server" => 54,
            "Virtual Machine" => 55,
            "Client Station" => 56,
            "Install Server" => 57,
            "ARCA DACS" => 58,
            "Packet Data Gateway" => 59,
            "RNG" => 60,
            "ADSP" => 61,
            "AP" => 62,
            "Firewall" => 63,
            "IDF" => 64,
            "MDF" => 65,
            "NX" => 66,
            "RFS" => 67,
            "UPS" => 68,
            "IPDU" => 69,
            "Device Config Server" => 70,
            "Trap Forwarder" => 71,
            "Jump Server" => 72,
            "ESX" => 73,
            "Gateway" => 74,
            "EXINDA" => 75,
            "Licensing Service" => 76,
            "Netcool Server" => 77,
            "DNS" => 78,
            "CPG" => 79,
            "MME" => 80,
            "SPM" => 81,
            "HSS" => 82,
            "PTT" => 83,
            "Security" => 84,
            "Object Server" => 85,
            "Firewall Bridge" => 86,
            "WebGUI" => 87,
            "Probe" => 88,
            "Impact" => 89,
            "Probe Server" => 90,
            "Guest WIFI" => 91,
            "OSS" => 92,
            "Base Radio" => 93,
            "Short Data Router" => 94,
            "Telephony" => 95,
            "AUC" => 96,
            "OSP" => 97,
            "Core" => 98,
            "OMADM" => 99,
            "Unknown" => 100,
            "MUX" => 101,
            "CCE" => 102,
            "Rectifier" => 103,
            "Alias Server" => 105,
            "Core Dispatch Comm Server" => 107,
            "MTIG" => 109,
            _ => 100
           
        };

        static int MapCiState(string state) => state switch
        {
            "Production" => 1,
            "PreProduction" => 2,
            "Decommissioned" => 3,
            _ => 1
        };

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.InitialDirectory = "C:";

            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                fullPath = openFileDialog.FileName;
                FilePathTextBox.Text = fullPath;
                Log($"Selected file: {fullPath}");
            }
        }

        private void Log(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"{DateTime.Now}: {message}{Environment.NewLine}");
                LogTextBox.ScrollToEnd();
            });
        }

        private void LogError(string message, Exception ex)
        {
            Log($"{message} {ex.Message}");
        }
    }
}
