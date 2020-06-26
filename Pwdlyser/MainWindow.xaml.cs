using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Diagnostics;
using System.Threading.Tasks;
using LiveCharts;
using LiveCharts.Wpf;
using System.Text;
using System.Net;
using ActiveDirectory;
using System.DirectoryServices;
using System.Xml.Linq;
using System.Windows.Navigation;

namespace Pwdlyser
{
    public partial class PwdlyserXMLDomainPromptWindow : Window1
    {
        public PwdlyserXMLDomainPromptWindow()
        {
            InitializeComponent();
        }
    }

    public partial class MainWindow : Window
    {
        // Define global vars 
        public string ImportFilename { get; internal set; }
        public string ImportFilePath { get; internal set; } // Password file import path
        public string WordlistImportFilename { get; internal set; } // Wordlist file import name
        public string[] ArrayPassList;
        public bool UserPassFileLoaded { get; internal set; }
        public string[] CommonPassListArray { get; internal set; }
        public string StringSummaryXMLOutput { get; internal set; }

        // Common/Dictionary Base-word lists
        internal string langdict = "";
        internal string[] langdictarray;

        internal string langdictlarge = "";
        internal string[] langdictlargearray;

        // Create datatables for analysis storage
        internal DataTable RawDataTable;
        internal DataTable RawDataTableBackup;
        internal DataTable DataTableCommon;
        internal DataTable DataTableCommonChart;
        internal DataTable DataTableCharacterType;
        internal DataTable DataTableCharacterTypeTemp;
        internal DataTable DataTableCharacter;
        internal DataTable DataTableCharacterTemp;
        internal DataTable DataTableCharacterFrequency;
        internal DataTable DataTableKeyboard;
        internal DataTable DataTableUsernames;
        internal DataTable DataTableReuse;
        internal DataTable DataTableReuseFreq;
        internal DataTable DataTableKeyboardFreq;
        internal DataTable DataTableDateFreq;
        internal DataTable DataTableDateTime;
        internal DataTable DataTableCompany;
        internal DataTable DataTableAdmin;
        internal DataTable DataTableHistoryAll;
        internal DataTable DataTableLength;
        internal DataTable DataTableWordList;
        internal DataTable DataTableWordListDistinct;
        internal DataTable DataTableHistoryReuse;
        internal DataTable DataTablePwned;
        internal DataTable DataTableWeak;
        internal DataTable ADUsers;
        internal DataTable DataTableSearch;
        internal DataTable DataTableReuseAdmin;
        internal DataTable DataTablePwnedAdmin;
        internal DataTable DataTableCharPosition;
        internal DataTable DataTableCharPositionTemp;
        internal DataTable DataTableFrequency;
        internal DataTable DataTableLengthFreq;
        internal DataTable DataTableTrailingSuffix;
        internal DataTable DataTableTrailingSuffixTemp;
        internal DataTable DataTableTrailingSuffixMask;
        internal DataTable ADAdmins;
        internal DataTable DataTableHashcatMasks;

        // Chart collections
        public SeriesCollection ChartTotalCrackedSeriesCollection { get; set; }
        public string[] ChartTotalCrackedLabels { get; set; }
        public Func<int, string> ChartTotalCrackedFormatter { get; set; }
        public SeriesCollection ChartAdminSeriesCollection { get; set; }
        public string[] ChartAdminLabels { get; set; }
        public Func<int, string> ChartAdminFormatter { get; set; }
        public SeriesCollection ChartWeakSeriesCollection { get; set; }
        public string[] ChartWeakLabels { get; set; }
        public Func<int, string> ChartWeakFormatter { get; set; }
        public SeriesCollection ChartCompanySeriesCollection { get; set; }
        public string[] ChartCompanyLabels { get; set; }
        public Func<int, string> ChartCompanyFormatter { get; set; }
        public SeriesCollection ChartUsernameSeriesCollection { get; set; }
        public string[] ChartUsernameLabels { get; set; }
        public Func<int, string> ChartUsernameFormatter { get; set; }
        public SeriesCollection ChartDateSeriesCollection { get; set; }
        public string[] ChartDateLabels { get; set; }
        public Func<int, string> ChartDateFormatter { get; set; }
        public SeriesCollection ChartKeyboardSeriesCollection { get; set; }
        public string[] ChartKeyboardLabels { get; set; }
        public Func<int, string> ChartKeyboardFormatter { get; set; }
        public SeriesCollection ChartReuseSeriesCollection { get; set; }
        public string[] ChartReuseLabels { get; set; }
        public Func<int, string> ChartReuseFormatter { get; set; }
        public SeriesCollection ChartCharacterSeriesCollection { get; set; }
        public string[] ChartCharacterLabels { get; set; }
        public Func<int, string> ChartCharacterFormatter { get; set; }
        public SeriesCollection ChartPwnedSeriesCollection { get; set; }
        public string[] ChartPwnedLabels { get; set; }
        public Func<int, string> ChartPwnedFormatter { get; set; }
        public SeriesCollection ChartLengthSeriesCollection { get; set; }
        public string[] ChartLengthLabels { get; set; }
        public Func<int, string> ChartLengthFormatter { get; set; }
        public SeriesCollection ChartFrequencySeriesCollection { get; set; }
        public string[] ChartFrequencyLabels { get; set; }
        public Func<int, string> ChartFrequencyFormatter { get; set; }
        public SeriesCollection ChartCommonSeriesCollection { get; set; }
        public string[] ChartCommonLabels { get; set; }
        public Func<int, string> ChartCommonFormatter { get; set; }

        // Common password check - dictionary specified
        public bool commondictcheck = false;
        public bool commondictchecklarge = false;

        // Pwned Passwords
        public bool PwnedAPIErrorResponse { get; private set; }
        public int maxpwncheck = 0;
        internal bool PwnedAdminCheckBool = false;

        // time Elapsed Strings
        internal string TimeElapsedCharacter;
        internal string TimeElapsedFrequency;
        internal string TimeElapsedLength;
        internal string TimeElapsedCommonDict;
        internal string TimeElapsedReuse;
        internal string TimeElapsedReuseAdmin;
        internal string TimeElapsedKeyboard;
        internal string TimeElapsedDateTime;
        internal string TimeElapsedCompany;
        internal string TimeElapsedAdmin;
        internal string TimeElapsedUsername;
        internal string TimeElapsedHistory;
        internal string TimeElapsedWeak;
        internal string TimeElapsedPwned;
        internal string TimeElapsedHashcat;
        internal string TimeElapsedTotalRun;

        // Summary Analysis Bool Checks
        public bool BoolSummaryAdmin = false;
        public bool BoolSummaryLength = false;
        public bool BoolSummaryCompany = false;
        public bool BoolSummaryCommon = false;
        public bool BoolSummaryUser = false;
        public bool BoolSummaryKeyboard = false;
        public bool BoolSummaryReuse = false;
        public bool BoolSummaryReuseAdmin = false;
        public bool BoolIncludeHistory = false;
        public bool BoolSummaryDate = false;
        public bool BoolSummaryHistory = false;
        public bool BoolSummaryCharacter = false;
        public bool BoolSummaryPwned = false;
        public bool BoolSummaryWeak = false;
        public string StringCharacterType = "";
        public bool ContinueSummaryXMLGeneration = false;

        // Plaintext summary string output variables
        public string ExecutiveSummaryContent = "";
        public string TechnicalSummaryContent = "";
        public string TechnicalSummaryContentMD = "";
        public string RecommendSummaryContent = "";

        internal List<string> list = new List<string>(); // No idea what this is for - wordlist generator?

        public MainWindow()
        {
            InitializeComponent();
            main();
        }

        public void main()
        {
            Dispatcher.Invoke(() => GetDomainName());

            CreateRawDataTable();
            CreateADUserTable();

            TabAnalysis.Visibility = Visibility.Collapsed;
            TabWordlist.Visibility = Visibility.Collapsed;
            TabSummary.Visibility = Visibility.Collapsed;
            TabXMLOutput.Visibility = Visibility.Collapsed;

            if (Properties.Settings.Default.GetStartedAccepted == true)
            {
                TabItemDashboard.Visibility = Visibility.Collapsed;
                TabAnalysis.Visibility = Visibility.Visible;
                TabWordlist.Visibility = Visibility.Visible;
                TabXMLOutput.Visibility = Visibility.Visible;
                //TabXMLOutput.Visibility = Visibility.Collapsed;
                TabSummary.Visibility = Visibility.Collapsed;
                Dispatcher.BeginInvoke((Action)(() => MainTabControl.SelectedIndex = 1));
                Dispatcher.BeginInvoke((Action)(() => AnalysisTabControl.SelectedIndex = 0));
                Dispatcher.Invoke(() => LoadXMLSummarySettings());
            }

            Properties.Settings.Default.SettingTemp = "1";
            Properties.Settings.Default.Save();

            SetLangDictResource();

            TextBoxLengthMinimum.Text = Properties.Settings.Default.LengthMinimumUserSetting;

            SettingsAddAdmins();
            SettingsAddCompanies();

            TextBoxTotalManualAddDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

            try
            {


                LoadTotalXMLFile();
            }
            catch
            {
                var filePath = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.TotalOverTimeSettingsPath);
                if (!File.Exists(filePath))
                {

                    CreateTotalXML();


                }
                LoadTotalXMLFile();

            }
        }

        public void GetDomainName()
        {
            try
            {
                string CurrentADDomain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
                TextBoxDomainName.Text = CurrentADDomain;
                SummaryXMLDomain.Text = CurrentADDomain;
            }
            catch
            {
                return;
            }
        }

        private void ButtonSettingsEnableAll_Click(object sender, RoutedEventArgs e)
        {
            CheckAnalysisAdministrator.IsChecked = true;
            CheckAnalysisCharacter.IsChecked = true;
            CheckAnalysisCommon.IsChecked = true;
            CheckAnalysisCompany.IsChecked = true;
            CheckAnalysisDate.IsChecked = true;
            CheckAnalysisFrequency.IsChecked = true;
            CheckAnalysisHashcat.IsChecked = true;
            CheckAnalysisHistory.IsChecked = true;
            CheckAnalysisKeyboard.IsChecked = true;
            CheckAnalysisLength.IsChecked = true;
            CheckAnalysisReuse.IsChecked = true;
            CheckAnalysisUsername.IsChecked = true;
            CheckAnalysisPwned.IsChecked = true;
            CheckAnalysisWeak.IsChecked = true;


            TabCharacterAnalysis.Visibility = Visibility.Visible;
            TabLengthAnalysis.Visibility = Visibility.Visible;
            TabFrequencyAnalysis.Visibility = Visibility.Visible;
            TabBaseAnalysis.Visibility = Visibility.Visible;
            TabReuseAnalysis.Visibility = Visibility.Visible;
            TabKeyboardAnalysis.Visibility = Visibility.Visible;
            TabDateAnalysis.Visibility = Visibility.Visible;
            TabCompanyAnalysis.Visibility = Visibility.Visible;
            TabUsernameAnalysis.Visibility = Visibility.Visible;
            TabAdminAnalysis.Visibility = Visibility.Visible;
            TabHashcatAnalysis.Visibility = Visibility.Visible;
            TabHistoryAnalysis.Visibility = Visibility.Visible;
            TabPwnedPasswords.Visibility = Visibility.Visible;
            TabWeakAnalysis.Visibility = Visibility.Visible;
        }

        internal void UpdateStatusRecords()
        {
            Dispatcher.Invoke(() => LabelStatusLength.Text = (DataTableLength.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusAdmin.Text = (DataTableAdmin.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusCommon.Text = (DataTableCommon.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusFreq.Text = (DataTableFrequency.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusReuse.Text = (DataTableReuse.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusReuseAdmin.Text = (DataTableReuseAdmin.Rows.Count.ToString() + " Records"));

            Dispatcher.Invoke(() => LabelStatusDate.Text = (DataTableDateTime.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusCompany.Text = (DataTableCompany.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusHistory.Text = (DataTableHistoryReuse.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusWeak.Text = (DataTableWeak.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusHashcat.Text = (DataTableHashcatMasks.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusCharacter.Text = (DataTableCharacter.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusUsername.Text = (DataTableUsernames.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusKeyboard.Text = (DataTableKeyboard.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusPwned.Text = (DataTablePwned.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusFreq.Text = (DataTableFrequency.Rows.Count.ToString() + " Records"));
            Dispatcher.Invoke(() => LabelStatusTotalRun.Text = (RawDataTable.Rows.Count.ToString() + " Records"));
        }

        // Create Data Tables
        internal void CreateRawDataTable()
        {
            RawDataTable = new DataTable();
            RawDataTable.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                });

            RawDataTableBackup = new DataTable();
            RawDataTableBackup.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                });
        }

        internal void CreateWordlistDataTable()
        {
            DataTableWordList = new DataTable();
            DataTableWordList.Columns.AddRange(new[]
                                       {
                               new DataColumn("Passwords", typeof(string))
                });

            DataTableWordListDistinct = new DataTable();
            DataTableWordListDistinct.Columns.AddRange(new[]
                           {
                               new DataColumn("Passwords", typeof(string))
                });
        }

        internal void CreateADUserTable()
        {
            ADUsers = new DataTable();
            ADUsers.Columns.AddRange(new[]
                                       {
                               new DataColumn("Username", typeof(string)),
                               new DataColumn("Last Logon", typeof(int)),
                               new DataColumn("Days Since Last Set", typeof(int)),
                               new DataColumn("Account Status", typeof(string)),
                               new DataColumn("Description", typeof(string)),
                               new DataColumn("Cracked", typeof(bool)),

                });
            ADUsers.Columns["Cracked"].DefaultValue = false;
        }

        internal void CreateSearchTable()
        {
            DataTableSearch = new DataTable();
            DataTableSearch.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Hashcat Password Mask", typeof(string)),
                });
        }

        internal void RefreshDataGrids()
        {
            DataGridLength.ItemsSource = DataTableLength.AsDataView();
            DataGridHistoryReuse.ItemsSource = DataTableHistoryReuse.AsDataView();
            DataGridAdmin.ItemsSource = DataTableAdmin.AsDataView();
            DataGridWeak.ItemsSource = DataTableWeak.AsDataView();
            DataGridCompany.ItemsSource = DataTableCompany.AsDataView();
            DataGridDateTime.ItemsSource = DataTableDateTime.AsDataView();
            DataGridReuseAdmin.ItemsSource = DataTableReuseAdmin.AsDataView();
            DataGridReuse.ItemsSource = DataTableReuse.AsDataView();
            DataGridUsernames.ItemsSource = DataTableUsernames.AsDataView();
            DataGridHashcat.ItemsSource = DataTableHashcatMasks.AsDataView();
            DataGridCommon.ItemsSource = DataTableCommon.AsDataView();
            DataGridHashcatSuffixRules.ItemsSource = DataTableTrailingSuffixMask.AsDataView();
            DataGridCharacterType.ItemsSource = DataTableCharacterType.AsDataView();
            DataGridCharacter.ItemsSource = DataTableCharacter.AsDataView();
            DataGridKeyboard.ItemsSource = DataTableKeyboard.AsDataView();
            DataGridPwned.ItemsSource = DataTablePwned.AsDataView();
            DataGridPwnedAdmin.ItemsSource = DataTablePwnedAdmin.AsDataView();
            DataGridTrailingSuffix.ItemsSource = DataTableTrailingSuffix.AsDataView();
            DataGridCharPosition.ItemsSource = DataTableCharPosition.AsDataView();
            DataGridFrequency.ItemsSource = DataTableFrequency.AsDataView();
        }

        internal void CreateDataTables()
        {

            DataTableLength = new DataTable();
            DataTableLength.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Length", typeof(int)),
                });

            DataTableHistoryReuse = new DataTable();
            DataTableHistoryReuse.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Current Passwords", typeof(string)),
                               new DataColumn("Re-use Occurrences", typeof(string)),
                });

            DataTableHistoryAll = new DataTable();
            DataTableHistoryAll.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Current Passwords", typeof(string)),
                });

            DataTableAdmin = new DataTable();
            DataTableAdmin.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                });

            DataTableWeak = new DataTable();
            DataTableWeak.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                });

            DataTableCompany = new DataTable();
            DataTableCompany.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Company Names", typeof(string)),
            });

            DataTableDateTime = new DataTable();
            DataTableDateTime.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Base Words", typeof(string)),
            });

            DataTableDateFreq = new DataTable();
            DataTableDateFreq.Columns.AddRange(new[]
                           {
                               new DataColumn("Date Base Words", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
            });

            DataTableKeyboardFreq = new DataTable();
            DataTableKeyboardFreq.Columns.AddRange(new[]
                           {
                               new DataColumn("Base Words", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
            });

            DataTableReuseFreq = new DataTable();
            DataTableReuseFreq.Columns.AddRange(new[]
                           {
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
            });

            DataTableReuse = new DataTable();
            DataTableReuse.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Shared Accounts", typeof(string)),
            });

            DataTableReuseAdmin = new DataTable();
            DataTableReuseAdmin.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Admin Accounts", typeof(string)),
            });

            DataTableHashcatMasks = new DataTable();
            DataTableHashcatMasks.Columns.AddRange(new[]
                                       {
                               new DataColumn("Hashcat Mask", typeof(string)),
                               new DataColumn("Occurences", typeof(int)),
            });

            DataTableUsernames = new DataTable();
            DataTableUsernames.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
            });

            DataTableCommon = new DataTable();
            DataTableCommon.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Base Words", typeof(string)),
            });

            DataTableCommon.PrimaryKey = new DataColumn[] { DataTableCommon.Columns[0], DataTableCommon.Columns[1] };

            DataTableTrailingSuffixMask = new DataTable();
            DataTableTrailingSuffixMask.Columns.AddRange(new[]
                                       {
                               new DataColumn("Hashcat Rules", typeof(string)),
                               new DataColumn("Occurrences", typeof(int)),
            });


            DataTableCommonChart = new DataTable();
            DataTableCommonChart.Columns.AddRange(new[]
                           {
                               new DataColumn("Base Words", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
            });

            DataTableCharacterTypeTemp = new DataTable();
            DataTableCharacterTypeTemp.Columns.AddRange(new[]
                           {
                               new DataColumn("Character Type", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
            });

            DataTableCharacterTypeTemp.Columns["Occurrences"].DataType = Type.GetType("System.Int32");

            DataTableCharacterType = new DataTable();
            DataTableCharacterType.Columns.AddRange(new[]
                           {
                               new DataColumn("Character Type", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
            });
            DataTableCharacterType.Columns["Occurrences"].DataType = Type.GetType("System.Int32");


            DataTableCharacter = new DataTable();
            DataTableCharacter.Columns.AddRange(new[]
                                       {
                               new DataColumn("Characters", typeof(string)),
                               new DataColumn("Occurrences", typeof(int)),
            });

            DataTableCharacter.Columns["Occurrences"].DataType = Type.GetType("System.Int32");

            DataTableCharacterTemp = new DataTable();
            DataTableCharacterTemp.Columns.AddRange(new[]
                                       {
                               new DataColumn("Characters", typeof(string)),
                               new DataColumn("Occurrences", typeof(int)),
            });

            DataTableCharacterTemp.Columns["Occurrences"].DataType = Type.GetType("System.Int32");

            DataTableCharacterFrequency = new DataTable();
            DataTableCharacterFrequency.Columns.AddRange(new[]
                                       {
                               new DataColumn("Characters", typeof(string)),
                               new DataColumn("Character Type", typeof(string)),
            });

            DataTableKeyboard = new DataTable();
            DataTableKeyboard.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Keyboard Pattern", typeof(string)),
            });

            DataTablePwned = new DataTable();
            DataTablePwned.Columns.AddRange(new[]
                                       {
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Occurrences", typeof(int)),
            });

            DataTablePwnedAdmin = new DataTable();
            DataTablePwnedAdmin.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Occurrences", typeof(int)),
            });

            DataTableTrailingSuffix = new DataTable();
            DataTableTrailingSuffix.Columns.AddRange(new[]
                                       {
                               new DataColumn("Suffix", typeof(string)),
                               new DataColumn("Occurrences", typeof(int)),
            });

            DataTableTrailingSuffixTemp = new DataTable();
            DataTableTrailingSuffixTemp.Columns.AddRange(new[]
                                       {
                               new DataColumn("Suffix", typeof(string)),
            });

            DataTableSearch = new DataTable();
            DataTableSearch.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
            });

            DataTableCharPosition = new DataTable();
            DataTableCharPosition.Columns.AddRange(new[]
                                       {
                               new DataColumn("Position", typeof(string)),
                               new DataColumn("Character", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
            });

            DataTableCharPosition.Columns["Occurrences"].AllowDBNull = true;

            DataTableFrequency = new DataTable();
            DataTableFrequency.Columns.AddRange(new[]
                                           {
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Occurrences", typeof(int)),
            });

            DataTableFrequency.Columns["Occurrences"].DataType = Type.GetType("System.Int32");

            DataTableLengthFreq = new DataTable();
            DataTableLengthFreq.Columns.AddRange(new[]
                   {
                               new DataColumn("Length", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),

            });

            DataTableCharPositionTemp = new DataTable();
            DataTableCharPositionTemp.Columns.AddRange(new[]
                                       {
                               new DataColumn("1", typeof(string)),
                               new DataColumn("2", typeof(string)),
                               new DataColumn("3", typeof(string)),
                               new DataColumn("4", typeof(string)),
                               new DataColumn("5", typeof(string)),
                               new DataColumn("6", typeof(string)),
                               new DataColumn("7", typeof(string)),
                               new DataColumn("8", typeof(string)),
                               new DataColumn("9", typeof(string)),
                               new DataColumn("10", typeof(string)),
                               new DataColumn("11", typeof(string)),
                               new DataColumn("12", typeof(string)),
                               new DataColumn("13", typeof(string)),
                               new DataColumn("14", typeof(string)),
                               new DataColumn("15", typeof(string)),
                               new DataColumn("16", typeof(string)),
                               new DataColumn("17", typeof(string)),
                               new DataColumn("18", typeof(string)),
                               new DataColumn("19", typeof(string)),
                               new DataColumn("20", typeof(string)),
                               new DataColumn("21", typeof(string)),
                               new DataColumn("22", typeof(string)),
                               new DataColumn("23", typeof(string)),
                               new DataColumn("24", typeof(string)),
                               new DataColumn("25", typeof(string)),
                               new DataColumn("26", typeof(string)),
                               new DataColumn("27", typeof(string)),
                               new DataColumn("28", typeof(string)),
                               new DataColumn("29", typeof(string)),
                               new DataColumn("30", typeof(string)),
                               new DataColumn("31", typeof(string)),
                               new DataColumn("32", typeof(string)),
                               new DataColumn("33", typeof(string)),
                               new DataColumn("34", typeof(string)),
                               new DataColumn("35", typeof(string)),
                               new DataColumn("36", typeof(string)),
                               new DataColumn("37", typeof(string)),
                               new DataColumn("38", typeof(string)),
                               new DataColumn("39", typeof(string)),
                               new DataColumn("40", typeof(string)),
                               new DataColumn("41", typeof(string)),
                               new DataColumn("42", typeof(string)),

            });
            RefreshDataGrids(); // Sets initial binding after creation.
        }

        /*public void GenerateCharts()
        {
            Dispatcher.Invoke(() => GenerateChartFrequency());
            Dispatcher.Invoke(() => GenerateChartLength());
            Dispatcher.Invoke(() => GenerateChartCommon());
            Dispatcher.Invoke(() => GenerateChartCharacter());
            Dispatcher.Invoke(() => GenerateChartReuse());
            Dispatcher.Invoke(() => GenerateChartKeyboard());
            Dispatcher.Invoke(() => GenerateChartDate());
            Dispatcher.Invoke(() => GenerateChartUsername());
            Dispatcher.Invoke(() => GenerateChartCompany());
            Dispatcher.Invoke(() => GenerateChartAdmin());
            Dispatcher.Invoke(() => GenerateChartPwned());
        }*/

        // Charts / Graphs
        public void GenerateTotalOverTime()
        {
            Dispatcher.Invoke(() => GenerateChartTotalCracked());
        }

        public void GenerateChartTotalCracked()
        {

            ChartTotalCracked.Visibility = Visibility.Visible;
            try
            {
                ChartTotalCracked.AxisX.RemoveAt(0);
                ChartTotalCracked.AxisY.RemoveAt(0);
            }
            catch
            {

            }

            var Collection = new ChartValues<int> { };
            int ChartLabelTotalCrackedCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRowView row in DataGridTotalCracked.Items)
            {
                LabelTemp.Add(row.Row[0].ToString());
                int test = int.Parse(row.Row[1].ToString());

                Collection.Add(test);
                ChartLabelTotalCrackedCount++;
            }


            ChartTotalCrackedSeriesCollection = new SeriesCollection
            {
                new LineSeries
                {
                    Title = "Percentage of Analysed Passwords Over Time",
                    Values = Collection,
                    LineSmoothness = 1
                }
            };

            ChartTotalCracked.AxisX.Add(new Axis
            {

                Labels = LabelTemp,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                },
                LabelsRotation = 45,
                Title = "Analysis Date"
            });



            ChartTotalCrackedFormatter = value => value.ToString("C");

            DataContext = this;
            ChartTotalCracked.Series = ChartTotalCrackedSeriesCollection;

            ChartTotalYAxis.MaxValue = 101;
            ChartTotalYAxis.Title = "Percentage of Total Passwords Obtained";
        }

        public void GenerateChartAdmin()
        {
            ChartAdminLabels = new[] { "" };

            ChartAdmin.Visibility = Visibility.Visible;
            ChartTitleAdmin.Visibility = Visibility.Visible;

            int totalPasswords = 0;
            int AdminPasswords = 0;

            AdminPasswords = DataTableAdmin.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
            totalPasswords = ListBoxAdmin.Items.Count;

            int otherPasswords = totalPasswords - AdminPasswords;

            Func<ChartPoint, string> labelPoint = chartPoint => string.Format("{0} ({1:P})", chartPoint.Y, chartPoint.Participation);

            ChartAdmin.Series = new SeriesCollection
            {
                new PieSeries
                {
                    Title = "Uncracked Administrator Accounts",
                    Values = new ChartValues<double> {otherPasswords},
                    DataLabels = true,
                    LabelPoint = labelPoint
                },

                new PieSeries
                {
                    Title = "Cracked Administrator Accounts",
                    Values = new ChartValues<double> {AdminPasswords},
                    PushOut = 15,
                    DataLabels = true,
                    LabelPoint = labelPoint

                }
            };

            ChartAdmin.LegendLocation = LegendLocation.Bottom;
            ProgressBarUpdate();
        }

        public void GenerateChartWeak()
        {
            ChartWeakLabels = new[] { "" };

            ChartWeak.Visibility = Visibility.Visible;
            ChartTitleWeak.Visibility = Visibility.Visible;

            int totalPasswords = 0;
            int WeakPasswords = 0;

            if (BoolIncludeHistory)
            {
                totalPasswords = int.Parse(RawDataTable.Rows.Count.ToString());
                WeakPasswords = int.Parse(DataTableWeak.Rows.Count.ToString());
            }
            else
            {
                totalPasswords = RawDataTable.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
                WeakPasswords = DataTableWeak.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
            }

            int otherPasswords = totalPasswords - WeakPasswords;

            Func<ChartPoint, string> labelPoint = chartPoint => string.Format("{0} ({1:P})", chartPoint.Y, chartPoint.Participation);

            ChartWeak.Series = new SeriesCollection
            {
                new PieSeries
                {
                    Title = "Other Passwords",
                    Values = new ChartValues<double> {otherPasswords},
                    DataLabels = true,
                    LabelPoint = labelPoint
                },

                new PieSeries
                {
                    Title = "Weak/Non-Compliant Passwords",
                    Values = new ChartValues<double> {WeakPasswords},
                    PushOut = 15,
                    DataLabels = true,
                    LabelPoint = labelPoint

                }
            };

            ChartWeak.LegendLocation = LegendLocation.Bottom;
            ProgressBarUpdate();
        }

        public void GenerateChartCompany()
        {
            ChartCompanyLabels = new[] { "" };

            ChartCompany.Visibility = Visibility.Visible;
            ChartTitleCompany.Visibility = Visibility.Visible;

            int totalPasswords = 0;
            int CompanyPasswords = 0;

            if (BoolIncludeHistory)
            {
                totalPasswords = int.Parse(RawDataTable.Rows.Count.ToString());
                CompanyPasswords = int.Parse(DataTableCompany.Rows.Count.ToString());
            }
            else
            {
                totalPasswords = RawDataTable.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
                CompanyPasswords = DataTableCompany.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
            }

            int otherPasswords = totalPasswords - CompanyPasswords;

            Func<ChartPoint, string> labelPoint = chartPoint => string.Format("{0} ({1:P})", chartPoint.Y, chartPoint.Participation);

            ChartCompany.Series = new SeriesCollection
            {
                new PieSeries
                {
                    Title = "Other Passwords",
                    Values = new ChartValues<double> {otherPasswords},
                    DataLabels = true,
                    LabelPoint = labelPoint
                },

                new PieSeries
                {
                    Title = "Passwords Containing Company Name(s)",
                    Values = new ChartValues<double> {CompanyPasswords},
                    PushOut = 15,
                    DataLabels = true,
                    LabelPoint = labelPoint

                }
            };

            ChartCompany.LegendLocation = LegendLocation.Bottom;
            ProgressBarUpdate();
        }

        public void GenerateChartUsername()
        {
            ChartUsernameLabels = new[] { "" };

            ChartUsername.Visibility = Visibility.Visible;
            ChartTitleUsername.Visibility = Visibility.Visible;

            int totalPasswords = 0;
            int usernamePasswords = 0;

            if (BoolIncludeHistory)
            {
                totalPasswords = int.Parse(RawDataTable.Rows.Count.ToString());
                usernamePasswords = int.Parse(DataTableUsernames.Rows.Count.ToString());
            }
            else
            {
                totalPasswords = RawDataTable.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
                usernamePasswords = DataTableUsernames.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
            }

            int otherPasswords = totalPasswords - usernamePasswords;

            Func<ChartPoint, string> labelPoint = chartPoint => string.Format("{0} ({1:P})", chartPoint.Y, chartPoint.Participation);

            ChartUsername.Series = new SeriesCollection
            {
                new PieSeries
                {
                    Title = "Other Passwords",
                    Values = new ChartValues<double> {otherPasswords},
                    DataLabels = true,
                    LabelPoint = labelPoint
                },

                new PieSeries
                {
                    Title = "Passwords Containing Usernames",
                    Values = new ChartValues<double> {usernamePasswords},
                    PushOut = 15,
                    DataLabels = true,
                    LabelPoint = labelPoint

                }
            };

            ChartUsername.LegendLocation = LegendLocation.Bottom;
            ProgressBarUpdate();
        }

        public void GenerateChartDate()
        {
            ChartDateLabels = new[] { "" };

            ChartDate.Visibility = Visibility.Visible;
            ChartTitleDate.Visibility = Visibility.Visible;
            try
            {
                ChartDate.AxisY.RemoveAt(0);
            }
            catch
            {

            }

            var Collection = new ChartValues<int> { };
            int ChartLabelCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRow row in DataTableDateFreq.Rows)
            {
                LabelTemp.Add(row.ItemArray[0].ToString());
                int test = int.Parse(row.ItemArray[1].ToString());

                Collection.Add(test);
                ChartLabelCount++;
            }

            ChartDate.Series = new SeriesCollection
            {
                new RowSeries
                {
                    Title = "",
                    Values = Collection
                }
            };

            ChartDateLabels = LabelTemp.ToArray();

            ChartDate.AxisY.Add(new Axis
            {

                Labels = LabelTemp,
                Position = AxisPosition.LeftBottom,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                }
            });

            ChartDateFormatter = value => value.ToString("N");

            DataContext = this;
            ProgressBarUpdate();
        }

        public void GenerateChartKeyboard()
        {
            ChartKeyboardLabels = new[] { "" };

            ChartKeyboard.Visibility = Visibility.Visible;
            ChartTitleKeyboard.Visibility = Visibility.Visible;
            try
            {
                ChartKeyboard.AxisY.RemoveAt(0);
            }
            catch
            {

            }

            var Collection = new ChartValues<int> { };
            int ChartLabelCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRow row in DataTableKeyboardFreq.Rows)
            {

                LabelTemp.Add(row.ItemArray[0].ToString());
                int test = int.Parse(row.ItemArray[1].ToString());

                Collection.Add(test);
                ChartLabelCount++;

            }

            ChartKeyboard.Series = new SeriesCollection
            {
                new RowSeries
                {
                    Title = "",
                    Values = Collection
                }
            };

            ChartKeyboardLabels = LabelTemp.ToArray();

            ChartKeyboard.AxisY.Add(new Axis
            {

                Labels = LabelTemp,
                Position = AxisPosition.LeftBottom,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                }

            });

            ChartKeyboardFormatter = value => value.ToString("N");

            DataContext = this;
            ProgressBarUpdate();
        }

        public void GenerateChartReuse()
        {
            ChartReuseLabels = new[] { "" };

            ChartReuse.Visibility = Visibility.Visible;
            ChartTitleReuse.Visibility = Visibility.Visible;
            try
            {
                ChartReuse.AxisY.RemoveAt(0);
            }
            catch
            {

            }


            var Collection = new ChartValues<int> { };
            int ChartLabelCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRow row in DataTableReuseFreq.Rows)
            {

                LabelTemp.Add(row.ItemArray[0].ToString());
                int test = int.Parse(row.ItemArray[1].ToString());

                Collection.Add(test);
                ChartLabelCount++;

            }

            ChartReuse.Series = new SeriesCollection
            {
                new RowSeries
                {
                    Title = "",
                    Values = Collection
                }
            };

            ChartReuseLabels = LabelTemp.ToArray();

            ChartReuse.AxisY.Add(new Axis
            {

                Labels = LabelTemp,
                Position = AxisPosition.LeftBottom,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                }

            });

            ChartReuseFormatter = value => value.ToString("N");

            DataContext = this;
            ProgressBarUpdate();
        }

        public void GenerateChartCharacter()
        {
            ChartCharacterLabels = new[] { "" };

            ChartCharacter.Visibility = Visibility.Visible;
            ChartTitleCharacter.Visibility = Visibility.Visible;


            Func<ChartPoint, string> labelPoint = chartPoint => string.Format("{0} ({1:P})", chartPoint.Y, chartPoint.Participation);

            ChartCharacter.Series = new SeriesCollection { };

            foreach (DataRow row in DataTableCharacterType.Rows)
            {
                string charsetName = row.ItemArray[0].ToString();
                int charsetCount = int.Parse(row.ItemArray[1].ToString());

                ChartCharacter.Series.Add(
                new PieSeries
                {
                    Title = charsetName,
                    Values = new ChartValues<double> { charsetCount },
                    DataLabels = true,
                    LabelPoint = labelPoint
                });
            };

            ChartCharacter.LegendLocation = LegendLocation.Bottom;
            ProgressBarUpdate();
        }

        public void GenerateChartPwned()
        {
            ChartPwnedLabels = new[] { "" };

            ChartPwned.Visibility = Visibility.Visible;
            ChartTitlePwned.Visibility = Visibility.Visible;
            try
            {
                ChartPwned.AxisY.RemoveAt(0);
            }
            catch
            {

            }

            //var Labels = new List<string> { };
            var Collection = new ChartValues<int> { };
            int ChartLabelCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRowView row in DataGridPwned.Items)
            {
                if (ChartLabelCount < 50)
                {
                    LabelTemp.Add(row.Row[0].ToString());
                    int test = int.Parse(row.Row[1].ToString());

                    Collection.Add(test);
                    ChartLabelCount++;
                }
                else
                {
                    break;
                }
            }

            ChartPwned.Series = new SeriesCollection
            {
                new RowSeries
                {
                    Title = "",
                    Values = Collection,

                }
            };


            ChartPwnedLabels = LabelTemp.ToArray();

            ChartPwned.AxisY.Add(new Axis
            {

                Labels = LabelTemp,
                Position = AxisPosition.LeftBottom,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                }

            });

            ChartPwnedFormatter = value => value.ToString("N");

            DataContext = this;
            ProgressBarUpdate();
        }

        public void GenerateChartLength()
        {
            ChartLengthLabels = new[] { "" };

            ChartLength.Visibility = Visibility.Visible;
            ChartTitleLength.Visibility = Visibility.Visible;
            try
            {
                ChartLength.AxisY.RemoveAt(0);
            }
            catch
            {

            }


            var Collection = new ChartValues<int> { };
            int ChartLabelCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRowView row in DataGridLengthFreq.Items)
            {
                if (ChartLabelCount < 10)
                {
                    LabelTemp.Add(row.Row[0].ToString());
                    int test = int.Parse(row.Row[1].ToString());

                    Collection.Add(test);
                    ChartLabelCount++;
                }
                else
                {
                    break;
                }
            }



            ChartLength.Series = new SeriesCollection
            {
                new RowSeries
                {
                    Title = "",
                    Values = Collection
                }
            };


            ChartLengthLabels = LabelTemp.ToArray();

            ChartLength.AxisY.Add(new Axis
            {

                Labels = LabelTemp,
                Position = AxisPosition.LeftBottom,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                }

            });

            ChartLengthFormatter = value => value.ToString("N");

            DataContext = this;
            ProgressBarUpdate();
        }

        public void GenerateChartFrequency()
        {
            ChartFrequencyLabels = new[] { "" };

            ChartFrequency.Visibility = Visibility.Visible;
            ChartTitleFrequency.Visibility = Visibility.Visible;
            try
            {
                ChartFrequency.AxisY.RemoveAt(0);
            }
            catch
            {

            }


            var Collection = new ChartValues<int> { };
            int ChartLabelFrequencyCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRowView row in DataGridFrequency.Items)
            {
                if (ChartLabelFrequencyCount < 10)
                {
                    LabelTemp.Add(row.Row[0].ToString());
                    int test = int.Parse(row.Row[1].ToString());

                    Collection.Add(test);
                    ChartLabelFrequencyCount++;
                }
                else
                {
                    break;
                }
            }

            ChartFrequency.Series = new SeriesCollection
            {
                new RowSeries
                {
                    Title = "",
                    Values = Collection
                }
            };

            ChartFrequencyLabels = LabelTemp.ToArray();

            ChartFrequency.AxisY.Add(new Axis
            {

                Labels = LabelTemp,
                Position = AxisPosition.LeftBottom,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                }

            });

            ChartFrequencyFormatter = value => value.ToString("N");

            DataContext = this;
            ProgressBarUpdate();
        }

        public void GenerateChartCommon()
        {
            ChartCommonLabels = new[] { "" };

            ChartCommon.Visibility = Visibility.Visible;
            ChartTitleCommon.Visibility = Visibility.Visible;
            try
            {
                ChartCommon.AxisX.RemoveAt(0);
            }
            catch
            {

            }

            var Collection = new ChartValues<int> { };
            int ChartLabelCount = 0;
            List<string> LabelTemp = new List<string> { };

            foreach (DataRow row in DataTableCommonChart.Rows)
            {
                if (ChartLabelCount < 10)
                {
                    LabelTemp.Add(row.ItemArray[0].ToString());
                    int test = int.Parse(row.ItemArray[1].ToString());

                    Collection.Add(test);
                    ChartLabelCount++;
                }
                else
                {
                    break;
                }
            }

            ChartCommon.Series = new SeriesCollection
            {
                new ColumnSeries
                {
                    Title = "",
                    Values = Collection
                }
            };

            ChartCommonLabels = LabelTemp.ToArray();

            ChartCommon.AxisX.Add(new Axis
            {

                Labels = LabelTemp,
                Position = AxisPosition.LeftBottom,
                Separator = new LiveCharts.Wpf.Separator // force the separator step to 1, so it always display all labels
                {
                    Step = 1,
                    IsEnabled = false //disable it to make it invisible.

                }
            });

            ChartCommonFormatter = value => value.ToString("N");

            DataContext = this;
            ProgressBarUpdate();
        }

        // Generate DataTables for Specific Charts
        public void ReuseAnalysisChart()
        {
            // Generate 'CommonPassOutput' for most common basewords
            // Used in Charts
            var q = DataTableReuse.Rows.OfType<DataRow>()
                .GroupBy(x => x[1].ToString())
                .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                .OrderByDescending(x => x.Count);

            if (DataTableReuseAdmin.Rows.Count > 0)
            {
                var a = DataTableReuseAdmin.Rows.OfType<DataRow>()
                    .GroupBy(x => x[1].ToString())
                    .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                    .OrderByDescending(x => x.Count);
                foreach (var z in a)
                {
                    DataTableReuseFreq.Rows.Add(z.Password, z.Count);
                }
            }

            foreach (var z in q)
            {
                DataTableReuseFreq.Rows.Add(z.Password, z.Count);
            }
            ProgressBarUpdate();
        }

        public void DateAnalysisChart()
        {

            // Generate 'CommonPassOutput' for most common basewords
            // Used in Charts
            var q = DataTableDateTime.Rows.OfType<DataRow>()
                .GroupBy(x => x[2].ToString())
                .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                .OrderByDescending(x => x.Count);

            foreach (var z in q)
            {
                DataTableDateFreq.Rows.Add(z.Password, z.Count);
            }
            ProgressBarUpdate();
        }

        public void KeyboardAnalysisChart()
        {

            // Generate 'CommonPassOutput' for most common basewords
            // Used in Charts
            var q = DataTableKeyboard.Rows.OfType<DataRow>()
                .GroupBy(x => x[2].ToString())
                .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                .OrderByDescending(x => x.Count);

            foreach (var z in q)
            {
                DataTableKeyboardFreq.Rows.Add(z.Password, z.Count);
            }
            ProgressBarUpdate();
        }

        public void CommonAnalysisChart()
        {
            // Generate 'CommonPassOutput' for most common basewords
            // Used in Charts
            // Sorted by occurrence
            var q = DataTableCommon.Rows.OfType<DataRow>()
                .GroupBy(x => x[2].ToString())
                .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                .OrderByDescending(x => x.Count);

            // Adds data to public datatable for charts
            foreach (var z in q)
            {
                DataTableCommonChart.Rows.Add(z.Password, z.Count);
            }
            ProgressBarUpdate();
        }

        private void AdminLoadNewFile()
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                DefaultExt = ".txt"
            };

            bool? result = dlg.ShowDialog();

            if (result.HasValue && result.Value)
            {
                var currentAdminUsers = ListBoxAdmin.Items;
                var adminUserImport = File.ReadAllLines(dlg.FileName);

                foreach (string admin in adminUserImport)
                {
                    currentAdminUsers.Add(admin);
                    Properties.Settings.Default.AdminAccountsList = admin + ";" + Properties.Settings.Default.AdminAccountsList;
                }
                Properties.Settings.Default.Save();
            }
        }

        private void SettingsAddAdmins()
        {
            string[] adminslist = (Properties.Settings.Default.AdminAccountsList.ToString()).Split(';');
            foreach (string admin in adminslist)
            {
                if (!ListBoxAdmin.Items.Contains(admin) && admin.Length > 0)
                {
                    ListBoxAdmin.Items.Add(admin);

                }
            }
        }

        private void SettingsAddCompanies()
        {
            string[] companylist = (Properties.Settings.Default.CompanyNameListSettings.ToString()).Split(';');
            foreach (string company in companylist)
            {
                if (!ListBoxCompany.Items.Contains(company) && company.Length > 0)
                {
                    ListBoxCompany.Items.Add(company);

                }
            }
        }

        public class WaitCursor : IDisposable
        {
            private Cursor _previousCursor;

            public WaitCursor()
            {
                _previousCursor = Mouse.OverrideCursor;

                Mouse.OverrideCursor = Cursors.Wait;
            }

            #region IDisposable Members

            public void Dispose()
            {
                Mouse.OverrideCursor = _previousCursor;
            }

            #endregion
        }

        /// Run Methods for Analysis

        private async Task RunReuse()
        {
            // Reuse / Sharing
            if (CheckAnalysisReuse.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusReuse.Text = ("Running"));
                Dispatcher.Invoke(() => LabelStatusReuseAdmin.Text = ("Running"));

                await Task.Run(() => AnalysisAllReuseSharing());
                await Task.Run(() => AnalysisAdminReuseSharing());
                await Task.Run(() => ReuseAnalysisChart());
                Dispatcher.Invoke(() => GenerateChartReuse());

            }
            else
            {
                DataGridReuse.ItemsSource = null;
                DataGridReuseAdmin.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusReuse.Text = ("Checks Not Enabled"));
                Dispatcher.Invoke(() => LabelStatusReuseAdmin.Text = ("Checks Not Enabled"));

            }
        }

        private async Task RunPwned()
        {
            // HIBP Analysis
            if (CheckAnalysisPwned.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusPwned.Text = ("Running"));

                await Task.Run(() => AnalysisPwned());
                Dispatcher.Invoke(() => GenerateChartPwned());
            }
            else
            {
                DataGridPwned.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusPwned.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunCompany()
        {
            // Organisation Name
            if (CheckAnalysisCompany.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusCompany.Text = ("Running"));
                await Task.Run(() => AnalysisCompanyName());
                Dispatcher.Invoke(() => GenerateChartCompany());

            }
            else
            {
                DataGridCompany.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusCompany.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunHistory()
        {
            if (CheckAnalysisHistory.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusHistory.Text = ("Running"));
                await Task.Run(() => AnalysisHistoryList());
            }
            else
            {
                DataGridHistoryReuse.ItemsSource = null;
                DataGridHistorySelect.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusHistory.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunCharacter()
        {
            // Character Frequency
            if (CheckAnalysisCharacter.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusCharacter.Text = ("Running"));
                await Task.Run(() => AnalysisCharacters());
                Dispatcher.Invoke(() => GenerateChartCharacter());

            }
            else
            {
                DataGridCharacter.ItemsSource = null;
                DataGridCharacterType.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusCharacter.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunLength()
#pragma warning restore CS1998 // Async method lacks 'await' operators and will run synchronously
        {
            // Password Length
            if (CheckAnalysisLength.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusLength.Text = ("Running"));
                Dispatcher.Invoke(() => AnalysisLength());
                Dispatcher.Invoke(() => AnalysisLengthFrequency());
                DataGridLengthFreq.ItemsSource = DataTableLengthFreq.AsDataView();
                Dispatcher.Invoke(() => GenerateChartLength());

            }
            else
            {
                DataGridLength.ItemsSource = null;
                DataGridLengthFreq.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusLength.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunAdmin()
        {
            if (CheckAnalysisAdministrator.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusAdmin.Text = ("Running"));
                await Task.Run(() => AnalysisAdminName());
                Dispatcher.Invoke(() => GenerateChartAdmin());

            }
            else
            {
                DataGridAdmin.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusAdmin.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunUsername()
        {
            // Usernames in Passwords
            if (CheckAnalysisUsername.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusUsername.Text = ("Running"));
                await Task.Run(() => AnalysisUsernameAsPass());
                Dispatcher.Invoke(() => GenerateChartUsername());

            }
            else
            {
                DataGridUsernames.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusUsername.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunDate()
        {
            // Date / Time Variation
            if (CheckAnalysisDate.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusDate.Text = ("Running"));

                await Task.Run(() => AnalysisDateTime());
                await Task.Run(() => DateAnalysisChart());
                Dispatcher.Invoke(() => GenerateChartDate());

            }
            else
            {
                DataGridDateTime.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusDate.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunKeyboard()
        {
            // Keyboard Walking
            if (CheckAnalysisKeyboard.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusKeyboard.Text = ("Running"));
                await Task.Run(() => AnalysisKeyboard());
                await Task.Run(() => KeyboardAnalysisChart());
                Dispatcher.Invoke(() => GenerateChartKeyboard());

            }
            else
            {
                DataGridKeyboard.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusKeyboard.Text = ("Checks Not Enabled"));
            }

        }

        private async Task RunHashcatMask()
        {
            // Hashcat Masks
            if (CheckAnalysisHashcat.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusHashcat.Text = ("Running"));
                await Task.Run(() => AnalysisHashcatMask());
            }
            else
            {
                DataGridHashcat.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusHashcat.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunCommonBase()
        {
            // Common Password Base Words
            if (CheckAnalysisCommon.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusCommon.Text = ("Running"));
                await Task.Run(() => AnalysisDictCommon());
                await Task.Run(() => CommonAnalysisChart());
                Dispatcher.Invoke(() => GenerateChartCommon());

            }
            else
            {
                DataGridCommon.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusCommon.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunFrequency()
        {
            // Password Frequency
            if (CheckAnalysisFrequency.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusFreq.Text = ("Running"));
                await Task.Run(() => AnalysisFrequency());
                Dispatcher.Invoke(() => GenerateChartFrequency());

            }
            else
            {
                DataGridFrequency.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusFreq.Text = ("Checks Not Enabled"));
            }
        }

        private async Task RunWeak()
        {
            if (CheckAnalysisWeak.IsChecked.Value)
            {
                Dispatcher.Invoke(() => LabelStatusWeak.Text = ("Running"));

                await Task.Run(() => AnalysisWeakPassword());
                Dispatcher.Invoke(() => GenerateChartWeak());
            }
            else
            {
                DataGridWeak.ItemsSource = null;
                Dispatcher.Invoke(() => LabelStatusWeak.Text = ("Checks Not Enabled"));
            }


        }

        private void ProgressBarUpdate()
        {
            Dispatcher.Invoke(() => analysisprogress.Value += 1);
        }

        private void ProgressBarReset()
        {
            Dispatcher.Invoke(() => analysisprogress.Value = 0);

        }

        private void ProgressBarMax()
        {

            int maxCount = 32;

            if (CheckAnalysisFrequency.IsChecked == false) { maxCount = maxCount - 3; } // 3
            if (CheckAnalysisLength.IsChecked == false) { maxCount = maxCount - 3; } // 3
            if (CheckAnalysisCharacter.IsChecked == false) { maxCount = maxCount - 4; } // 4
            if (CheckAnalysisCommon.IsChecked == false) { maxCount = maxCount - 2; } // 2
            // 11
            if (CheckAnalysisReuse.IsChecked == false) { maxCount = maxCount - 3; } // 2
            if (CheckAnalysisKeyboard.IsChecked == false) { maxCount = maxCount - 2; } // 2
            if (CheckAnalysisDate.IsChecked == false) { maxCount = maxCount - 2; }  // 2
            if (CheckAnalysisCompany.IsChecked == false) { maxCount = maxCount - 2; } // 2
            if (CheckAnalysisAdministrator.IsChecked == false) { maxCount = maxCount - 2; } // 2
            // 21
            if (CheckAnalysisUsername.IsChecked == false) { maxCount = maxCount - 2; } // 2
            if (CheckAnalysisHistory.IsChecked == false) { maxCount = maxCount - 1; } // 1
            if (CheckAnalysisWeak.IsChecked == false) { maxCount = maxCount - 2; } // 2
            if (CheckAnalysisPwned.IsChecked == false) { maxCount = maxCount - 3; } // 3
            if (CheckAnalysisHashcat.IsChecked == false) { maxCount = maxCount - 2; } // 2
            // 32

            analysisprogress.Maximum = maxCount;

        }

        public void SetTimeElapsedBinding()
        {
            Dispatcher.Invoke(() => LabelTimeCharacter.Text = TimeElapsedCharacter);
            Dispatcher.Invoke(() => LabelTimeFreq.Text = TimeElapsedFrequency);
            Dispatcher.Invoke(() => LabelTimeLength.Text = TimeElapsedLength);
            Dispatcher.Invoke(() => LabelTimeCommon.Text = TimeElapsedCommonDict);
            Dispatcher.Invoke(() => LabelTimeReuse.Text = TimeElapsedReuse);
            Dispatcher.Invoke(() => LabelTimeReuseAdmin.Text = TimeElapsedReuseAdmin);
            Dispatcher.Invoke(() => LabelTimeKeyboard.Text = TimeElapsedKeyboard);
            Dispatcher.Invoke(() => LabelTimeDate.Text = TimeElapsedDateTime);
            Dispatcher.Invoke(() => LabelTimeCompany.Text = TimeElapsedCompany);
            Dispatcher.Invoke(() => LabelTimeAdmin.Text = TimeElapsedAdmin);
            Dispatcher.Invoke(() => LabelTimeUsername.Text = TimeElapsedUsername);
            Dispatcher.Invoke(() => LabelTimeHistory.Text = TimeElapsedHistory);
            Dispatcher.Invoke(() => LabelTimeWeak.Text = TimeElapsedWeak);
            Dispatcher.Invoke(() => LabelTimePwned.Text = TimeElapsedPwned);
            Dispatcher.Invoke(() => LabelTimeHashcat.Text = TimeElapsedHashcat);

            Dispatcher.Invoke(() => LabelTimeTotalRun.Text = TimeElapsedTotalRun);
        }
       
        // Analysis Method
        private async Task RunAnalysisChecks()
        {
            TabSummary.Visibility = Visibility.Collapsed;
            TabXMLOutput.Visibility = Visibility.Collapsed;

            ProgressBarMax();
            ClearStatusAndTime();

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            ProgressBarReset();
            ButtonRunAnalysis.Content = "Please Wait";
            ButtonGenerateSummary.Content = "Please Wait";
            ButtonRunAnalysis.IsEnabled = false;
            ButtonGenerateSummary.IsEnabled = false;

            Dispatcher.Invoke(() => LabelStatusTotalRun.Text = ("Running"));

            // If enabled then history users are included in analysis results
            if (CheckIncludeHistory.IsChecked.Value)
            {
                BoolIncludeHistory = true;
            }
            else
            {
                BoolIncludeHistory = false;
            }

            CommonPassListArray = ListBoxBaseword.Items.OfType<string>().ToArray(); ;

            //DataGridReuse.ItemsSource = null;
            //DataGridReuseAdmin.ItemsSource = null;
            DataTableReuse.Clear();
            DataTableReuseAdmin.Clear();

            //DataGridReuse.Items.Clear();
            //DataGridReuseAdmin.Items.Clear();

            Dispatcher.Invoke(() => SetMaxPwnedCheck());
            Dispatcher.Invoke(() => CheckAdminPwnedBoolSet());

            await Task.WhenAll(AnalysisPasswordSuffix(), RunDate(), RunKeyboard(), RunCompany(),
                RunCommonBase(), RunFrequency(), RunReuse(), RunHashcatMask(), RunPwned(),
                RunHistory(), RunCharacter(), RunLength(), RunAdmin(), RunUsername(),
                RunWeak(), AnalysisCharPosition());


            await PutTaskDelay();

            if (PwnedAdminCheckBool == true)
            {
                AnalysisPwnedAdmin();
            }


            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedTotalRun = elapsedTime.ToString();

            RefreshDataGrids();
            UpdateStatusRecords();
            SetTimeElapsedBinding();

            ButtonRunAnalysis.Content = "Run Analysis";
            ButtonGenerateSummary.Content = "Generate Summary";
            ButtonRunAnalysis.IsEnabled = true;
            ButtonGenerateSummary.IsEnabled = true;

            TabSummary.Visibility = Visibility.Visible;
            TabXMLOutput.Visibility = Visibility.Visible;

            if (DataGridUserProperties.Items.Count >= RawDataTable.Rows.Count)
            {
                AddToProjectXML();
            }

            Dispatcher.Invoke(() => CreateWordlistAnalysis.IsEnabled = true);

            Dispatcher.Invoke(() => GenerateTotalOverTime());
        }

        // Analysis Section
        private async Task AnalysisCharPosition()
        {
            if (CheckAnalysisCharacter.IsChecked.Value)
            {
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();


                var lower_regex = new Regex(@"[a-z]");
                var higher_regex = new Regex(@"[A-Z]");
                var numeric_regex = new Regex(@"[0-9]");
                var special_regex = new Regex(@"[!@£$%^&*()\[\]:;\\\/<>{}]");
                var history_regex = new Regex(@"_history[0-9]");

                int maxLength = 0;

                foreach (DataRow row in RawDataTable.Rows)
                {
                    // Set from raw data
                    string username = row[0].ToString();
                    string password = row[1].ToString();

                    if (password.Length > maxLength)
                    {
                        maxLength = password.Length;
                    }

                    if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                    {
                        continue;
                    }
                    DataRow temprow = DataTableCharPositionTemp.NewRow();

                    for (int i = 0; i < password.Length; i++)
                    {
                        string c = password[i].ToString();
                        if (c == null)
                        {
                            c = "";
                        }
                        temprow[(i + 1).ToString()] = c.ToString();
                    }
                    DataTableCharPositionTemp.Rows.Add(temprow);
                }

                int selectColumn = 0;
                for (int i = 0; i < maxLength; i++)
                {
                    try
                    {


                        var q = DataTableCharPositionTemp.Rows
                            .Cast<DataRow>()
                            .GroupBy(x => x.Field<string>(selectColumn))
                            .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                            .OrderByDescending(x => x.Count);

                        int linqQueryNull = 0;
                        var a = q.First();
                        if (a.Password == null)
                        {
                            foreach (var z in q)
                            {
                                if (z.Password != null)
                                {
                                    a = q.ElementAt(linqQueryNull);
                                    break;
                                }
                                linqQueryNull++;
                            }
                        }
                        DataTableCharPosition.Rows.Add((selectColumn + 1).ToString(), a.Password.ToString(), a.Count.ToString());
                    }

                    catch (Exception ex)
                    {
                        Console.Write(ex);
                    }

                    selectColumn++;
                }

                stopWatch.Stop();
                TimeSpan ts = stopWatch.Elapsed;

                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                    ts.Hours, ts.Minutes, ts.Seconds,
                    ts.Milliseconds / 10);

            }
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisAdminName()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");

            foreach (DataRow row in RawDataTable.Rows)
            {
                if (ListBoxAdmin.Items.Count > 0)
                {
                    string username = row[0].ToString();
                    string tempUsername = username;
                    string password = row[1].ToString();

                    foreach (string Admin in ListBoxAdmin.Items)
                    {
                        string tempAdmin = Admin;
                        if (username.Contains('\\') && !Admin.Contains('\\'))
                        {
                            tempUsername = username.Split('\\')[1];
                        }
                        else if (!username.Contains('\\') && Admin.Contains('\\'))
                        {
                            tempAdmin = Admin.Split('\\')[1];
                        }
                        if (tempAdmin.ToLower().Equals(tempUsername.ToLower()))
                        {
                            if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                            {
                                continue;
                            }

                            DataTableAdmin.Rows.Add(username, password);
                            count++;
                            if (BoolSummaryAdmin == false)
                            {
                                BoolSummaryAdmin = true;
                            }
                        }
                    }
                }
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);


            TimeElapsedAdmin = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        private void AnalysisPwnedAdmin()
        {
            PwnedAPIErrorResponse = false;

            string ConnectivityCheck = PwnedAPIRequest("ABCDE");

            if (ConnectivityCheck == "ERROR")
            {
                return;
            }

            Parallel.ForEach(DataTableAdmin.AsEnumerable(), new ParallelOptions { MaxDegreeOfParallelism = 1 }, row => //
            {

                string password = row[1].ToString();
                string username = row[0].ToString();

                string sha1Hash = Sha1Hash(password);
                if (PwnedAPIErrorResponse == false)
                {
                    string apiResponse = PwnedAPIRequest(sha1Hash.Substring(0, 5)).ToString();
                    if (apiResponse == "ERROR")
                    {
                        PwnedAPIErrorResponse = true;

                        return;
                    }

                    string[] apiResponseArray = (apiResponse.Replace("\r\n", "\n")).Split("\n".ToCharArray());

                    foreach (string hash in apiResponseArray)
                    {

                        if ((sha1Hash.Substring(0, 5) + hash.Split(':')[0].ToLower()).Contains(sha1Hash.ToLower()))
                        {
                            int occurrences = int.Parse(hash.Split(':')[1]);
                            DataRow rowAdd = DataTablePwnedAdmin.NewRow();
                            rowAdd[0] = username;
                            rowAdd[1] = password;
                            rowAdd[2] = occurrences;
                            DataTablePwnedAdmin.Rows.Add(rowAdd);
                        }
                    }

                }
                else
                {
                    return;
                }
            });

            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisFrequency()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");

            DataTable DataTableFreqTemp = new DataTable();
            DataTableFreqTemp.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                });

            //Parallel.ForEach(RawDataTable.Rows.Cast<DataRow>(), new ParallelOptions { MaxDegreeOfParallelism = 8 }, row =>
            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data
                string username = row[0].ToString();
                string password = row[1].ToString();

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                }
                DataTableFreqTemp.Rows.Add(username, password);
            };

            var q = DataTableFreqTemp.Rows.OfType<DataRow>()
                .GroupBy(x => x[1].ToString())
                .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                .OrderByDescending(x => x.Count);

            foreach (var z in q)
            {
                DataTableFrequency.Rows.Add(z.Password, z.Count);
                count++;
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedFrequency = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisSetDictSize()
        {
            if (CheckBoxBaseWordDict.IsChecked == true)
            {
                commondictcheck = true;
                commondictchecklarge = false;
                Dispatcher.Invoke(() => CheckBoxBaseWordDictLarge.IsChecked = false);
            }
            else if (CheckBoxBaseWordDictLarge.IsChecked == true)
            {
                commondictcheck = true;
                commondictchecklarge = true;
                Dispatcher.Invoke(() => CheckBoxBaseWordDict.IsChecked = false);

            }
            else if ((CheckBoxBaseWordDict.IsChecked == false) && (CheckBoxBaseWordDictLarge.IsChecked == false))
            {
                commondictcheck = false;
                commondictchecklarge = false;
                Dispatcher.Invoke(() => CheckBoxBaseWordDict.IsChecked = false);
                Dispatcher.Invoke(() => CheckBoxBaseWordDictLarge.IsChecked = false);

            }
        }

        public void AnalysisDictCommon()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");
            Dispatcher.Invoke(() => AnalysisSetDictSize());
            Dispatcher.Invoke(() => DataGridCommon.ItemsSource = null);

            //Parallel.ForEach(RawDataTable.AsEnumerable(), new ParallelOptions { MaxDegreeOfParallelism = 4 }, row =>
            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data
                string username = row[0].ToString();
                string password = row[1].ToString();
                string commonPassOutput = (password);

                if (password.Length > 0)
                {
                    commonPassOutput = CommonPass(password);

                }
                else
                {
                    return;
                }

                if (commonPassOutput.ToString().Equals("NOTVALID"))
                {
                    // Do nothing
                    continue;
                }
                else
                {
                    if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                    {
                        //return;
                        continue;
                    }
                    DataRow newRow = DataTableCommon.NewRow();
                    newRow[0] = username;
                    newRow[1] = password;
                    newRow[2] = commonPassOutput;
                    try
                    {

                        DataTableCommon.Rows.Add(newRow);
                    }

                    catch
                    {
                        continue;
                    }
                    count++;
                    if (BoolSummaryCommon == false)
                    {
                        BoolSummaryCommon = true;
                    }
                }
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedCommonDict = elapsedTime.ToString();
            ProgressBarUpdate();
        }

        public void AnalysisPwned()
        {
            PwnedAPIErrorResponse = false;

            string ConnectivityCheck = PwnedAPIRequest("ABCDE");

            if (ConnectivityCheck == "ERROR")
            {
                return;
            }

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");

            DataTable DataTablePwnedTemp = new DataTable();
            DataTablePwnedTemp.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                });

            DataTable DataTablePwnedOccurrence = new DataTable();
            DataTablePwnedOccurrence.Columns.AddRange(new[]
                                       {
                               new DataColumn("Passwords", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),
                });

            DataTablePwnedOccurrence.Columns["Occurrences"].DataType = Type.GetType("System.Int32");

            // Grabs raw data
            foreach (DataRow row in RawDataTable.Rows)
            {
                DataRow tmpRow = DataTablePwnedTemp.NewRow();

                string username = row[0].ToString();
                string password = row[1].ToString();

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue; //continue;
                }

                tmpRow[0] = username;
                tmpRow[1] = password;
                DataTablePwnedTemp.Rows.Add(tmpRow); // adds to temp datatable
            }

            // sorts datatable by most frequent (similar to frequency analysis)
            var q = DataTablePwnedTemp.Rows.OfType<DataRow>()
                .GroupBy(x => x[1].ToString())
                .Select(g => new { Password = g.Key, Count = g.Count(), Rows = g.ToList() })
                .OrderByDescending(x => x.Count);

            // adds to second temporary datatable
            foreach (var z in q)
            {
                DataTablePwnedOccurrence.Rows.Add(z.Password, z.Count);
                count++;
            }

            int topCount = 0;

            PwnedAPIErrorResponse = false;

            count = 0;

            Parallel.ForEach(DataTablePwnedOccurrence.AsEnumerable(), new ParallelOptions { MaxDegreeOfParallelism = 1 }, row => //
            //foreach (DataRow row in DataTablePwnedOccurrence.Rows)//Parallel.ForEach(distinctValues.AsEnumerable(), row =>
            {
                if (topCount < maxpwncheck) // top 50 passwords to check
                {
                    // Set from raw data
                    string password = row[0].ToString();

                    string sha1Hash = Sha1Hash(password);
                    if (PwnedAPIErrorResponse == false)
                    {
                        string apiResponse = PwnedAPIRequest(sha1Hash.Substring(0, 5)).ToString();
                        if (apiResponse == "ERROR")
                        {
                            PwnedAPIErrorResponse = true;

                            return;
                        }



                        string[] apiResponseArray = (apiResponse.Replace("\r\n", "\n")).Split("\n".ToCharArray());



                        foreach (string hash in apiResponseArray)
                        {

                            if ((sha1Hash.Substring(0, 5) + hash.Split(':')[0].ToLower()).Contains(sha1Hash.ToLower()))
                            {
                                int occurrences = int.Parse(hash.Split(':')[1]);
                                DataRow rowAdd = DataTablePwned.NewRow();
                                rowAdd[0] = password;
                                rowAdd[1] = occurrences;
                                DataTablePwned.Rows.Add(rowAdd);
                                topCount++;
                                count++;

                            }
                        }
                        if (BoolSummaryPwned == false)
                        {
                            BoolSummaryPwned = true;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
            });


            DataView dv = DataTablePwned.DefaultView;

            dv.Sort = "Occurrences desc";
            DataTable sortedDT = dv.ToTable();
            DataTablePwned = sortedDT;

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);


            TimeElapsedPwned = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisKeyboard()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");

            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data
                string username = row[0].ToString();
                string password = row[1].ToString();
                string keyboardPassOutput = KeyboardPass(password);

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                }

                if (keyboardPassOutput.ToString().Equals("NOTVALID"))
                {

                }
                else
                {
                    DataTableKeyboard.Rows.Add(username, password, keyboardPassOutput);
                    count++;
                    if (BoolSummaryKeyboard == false)
                    {
                        BoolSummaryKeyboard = true;
                    }
                }
            }

            DataView dv = DataTableKeyboard.DefaultView;
            dv.Sort = "Usernames asc";
            DataTable sortedDT = dv.ToTable();
            //return 
            DataTableKeyboard = sortedDT;

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedKeyboard = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisUsernameAsPass()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;


            var history_regex = new Regex(@"_history[0-9]");

            
            Parallel.ForEach(RawDataTable.Rows.Cast<DataRow>(), new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount * 10 }, (row) =>

            //foreach (DataRow row in RawDataTable.Rows)
            {

                // Set from raw data
                string username = row[0].ToString();
                string password = row[1].ToString();
                string tempUsername = username;
                if (username.Contains('\\'))
                {
                    tempUsername = tempUsername.ToLower().Split('\\')[1];
                }
                string deleeted = Deleet_password(password).ToString().ToLower();

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    return;
                    //return;
                }

                if (deleeted.ToLower().Contains(tempUsername.ToLower()))
                {
                    DataTableUsernames.Rows.Add(username, password);
                    count++;
                    if (BoolSummaryUser == false)
                    {
                        BoolSummaryUser = true;
                    }
                }
            });


            DataView dv = DataTableUsernames.DefaultView;
            dv.Sort = "Usernames asc";
            DataTable sortedDT = dv.ToTable();
            DataTableUsernames = sortedDT;

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);


            TimeElapsedUsername = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisAllReuseSharing()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            //stillWorking = true;
            int reuseCount = 0;

            var reuse_regex = new Regex(@"_history[0-9]");
            var reuse_regex_ext = new Regex(@"_history[0-9][0-9]");

            //foreach (DataRow row in RawDataTable.Rows)
            //Parallel.ForEach(RawDataTable.Rows.Cast<DataRow>(), new ParallelOptions { MaxDegreeOfParallelism = 4 }, (row) =>
            Parallel.ForEach(RawDataTable.Rows.Cast<DataRow>(), new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount * 10 }, (row) =>

            {
                string user1 = row[0].ToString();
                string pass1 = row[1].ToString();

                string user1temp;


                if (user1.Contains("\\"))
                {
                    user1temp = user1.Split('\\')[1];
                }
                else
                {
                    user1temp = user1;
                }

                if (user1temp.Length > 3)
                {

                    // Skips if user is _history
                    if (reuse_regex.Match(user1temp).Success)
                    {
                        return;
                    }

                    else
                    {
                        foreach (DataRow row_compare in RawDataTable.Rows) //user
                        {
                            string user2 = row_compare[0].ToString();
                            string pass2 = row_compare[1].ToString();

                            string user2temp;

                            if (user2.Contains("\\"))
                            {
                                user2temp = user2.Split('\\')[1];

                            }
                            else
                            {
                                user2temp = user2;
                            }

                            // Skips if compared user is _history
                            if (reuse_regex.Match(user2temp).Success)
                            {
                                continue;
                            }
                            else
                            {



                                if (user2temp.Length > 3)
                                {

                                    if ((user1temp.ToLower() != user2temp.ToLower()) && ((user1temp.ToLower().Contains(user2temp.ToLower())) && (pass1 == pass2)))
                                    {
                                        DataTableReuse.BeginLoadData();

                                        DataRow ReuseRow = DataTableReuse.NewRow();

                                        ReuseRow[0] = user2;
                                        ReuseRow[1] = pass1;
                                        ReuseRow[2] = user1;


                                        DataTableReuse.Rows.Add(ReuseRow);
                                        DataTableReuse.EndLoadData();
                                        //DataTableReuse.AcceptChanges();
                                        reuseCount++;
                                        if (BoolSummaryReuse == false)
                                        {
                                            BoolSummaryReuse = true;
                                        }
                                    }
                                }

                            }
                        }
                    }
                }
            });

            DataView dv = new DataView();
            dv = DataTableReuse.AsDataView();
            dv.Sort = "Usernames asc";
            DataTable sortedDT = dv.ToTable();

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            DataTableReuse = sortedDT;

            TimeElapsedReuse = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisAdminReuseSharing()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            int reuseCount = 0;

            var reuse_regex = new Regex(@"_history[0-9]");
            var reuse_regex_ext = new Regex(@"_history[0-9][0-9]");




            string[] adminlist = ListBoxAdmin.Items.OfType<string>().ToArray();

            if (adminlist.Count() > 0)
            {


                Parallel.ForEach(RawDataTable.Rows.Cast<DataRow>(), new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount * 10 }, (row) => // loop through raw data
                //foreach (DataRow row in RawDataTable.Rows)

                {
                    string user1temp;
                    string user1 = row[0].ToString();
                    string pass1 = row[1].ToString();

                    if (user1.Contains("\\"))
                    {
                        user1temp = user1.Split('\\')[1];
                    }
                    else
                    {
                        user1temp = user1;
                    }

                    // Skips if user is _history
                    if (reuse_regex.Match(user1temp).Success)
                    {
                        return;
                    }

                    else
                    {
                        foreach (DataRow row_compare in RawDataTable.Rows) // loop through admins
                        {
                            string user2 = row_compare[0].ToString();
                            string pass2 = row_compare[1].ToString();

                            string user2temp;

                            if (user2.Contains("\\"))
                            {
                                user2temp = user2.Split('\\')[1];

                            }
                            else
                            {
                                user2temp = user2;
                            }

                            // Skips if compared user is _history
                            if (reuse_regex.Match(user2temp).Success)
                            {
                                continue;
                            }
                            else
                            {


                                foreach (string adminuser in adminlist)
                                {

                                    if ((user1temp.ToLower() != user2temp.ToLower()) && ((user2temp.ToLower().Contains(adminuser.ToLower())) && (pass1 == pass2)))
                                    {
                                        DataRow AdminReuseRow = DataTableReuseAdmin.NewRow();

                                        AdminReuseRow[0] = user1;
                                        AdminReuseRow[1] = pass1;
                                        AdminReuseRow[2] = user2;

                                        DataTableReuseAdmin.Rows.Add(AdminReuseRow);
                                        reuseCount++;
                                        if (BoolSummaryReuseAdmin == false)
                                        {
                                            BoolSummaryReuseAdmin = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                });

                DataView dv = new DataView();
                dv = DataTableReuseAdmin.DefaultView;
                dv.Sort = "Usernames asc";
                DataTable sortedDT = dv.ToTable();

                DataTableReuseAdmin = sortedDT;

            }
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedReuseAdmin = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisLength()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");

            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data

                if (row[0] is null || row[1] is null)
                {
                    continue;
                }
                string username = row[0].ToString();
                string password = row[1].ToString();


                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                }

                if (password.Length < Int32.Parse(TextBoxLengthMinimum.Text))
                {
                    DataTableLength.Rows.Add(username, password, password.Length);
                    count++;
                    if (BoolSummaryLength == false)
                    {
                        BoolSummaryLength = true;
                    }
                }
            }

            DataTableLength.DefaultView.Sort = "Length asc";
            DataTableLength.DefaultView.ToTable();

            DataView dv = DataTableLength.DefaultView;
            dv.Sort = "Length asc";
            DataTable sortedDT = dv.ToTable();

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedLength = elapsedTime.ToString();
            DataTableLength = sortedDT.DefaultView.Table;
            SetTimeElapsedBinding();
        }

        public void AnalysisLengthFrequency()
        {
            var history_regex = new Regex(@"_history[0-9]");


            DataTable DataTableLengthFreqSorted = new DataTable();
            DataTableLengthFreqSorted.Columns.AddRange(new[]
               {
                               new DataColumn("Length", typeof(string)),
                });

            DataTable DataTableLengthFreqTemp = new DataTable();
            DataTableLengthFreqTemp.Columns.AddRange(new[]
               {
                               new DataColumn("Length", typeof(string)),
                               new DataColumn("Occurrences", typeof(string)),

                });

            DataTableLengthFreqTemp.Columns["Occurrences"].DataType = Type.GetType("System.Int32");

            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data

                if (row[0] is null || row[1] is null)
                {
                    continue;
                }
                string username = row[0].ToString();
                string password = row[1].ToString();


                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                }

                DataTableLengthFreqSorted.Rows.Add(password.Length);
                if (BoolSummaryLength == false)
                {
                    BoolSummaryLength = true;
                }
            }

            var query = from row in DataTableLengthFreqSorted.AsEnumerable()
                        group row by row.Field<string>("Length") into y
                        orderby y.Key
                        select new
                        {
                            Length = y.Key,
                            Count = y.Count()
                        };

            foreach (var z in query)
            {
                DataTableLengthFreqTemp.Rows.Add(z.Length, (z.Count));
            }

            DataView dvFreq = DataTableLengthFreqTemp.DefaultView;
            dvFreq.Sort = "Occurrences desc";
            DataTable sortedDTFreq = dvFreq.ToTable();


            DataTableLengthFreq = sortedDTFreq;
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisDateTime()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");


            foreach (DataRow row in RawDataTable.Rows)
            {

                // Set from raw data
                string username = row[0].ToString();
                string password = row[1].ToString();
                string dateTimeOutput = DateTimePass(password);

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                }

                if (dateTimeOutput.ToString() != "NOTVALID")
                {
                    DataTableDateTime.Rows.Add(username, password, dateTimeOutput);
                    count++;
                    if (BoolSummaryDate == false)
                    {

                        BoolSummaryDate = true;
                    }
                }
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedDateTime = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisCompanyName()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var history_regex = new Regex(@"_history[0-9]");

            foreach (DataRow row in RawDataTable.Rows)
            {
                if (ListBoxCompany.Items.Count > 0)
                {
                    string username = row[0].ToString();
                    string password = row[1].ToString();

                    if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                    {
                        continue;
                    }

                    foreach (string OrgName in ListBoxCompany.Items)
                    {
                        if (Deleet_password(password.ToLower()).Contains(Deleet_password(OrgName.ToLower())))
                        {
                            DataTableCompany.Rows.Add(username, password, OrgName);
                            count++;
                            if (BoolSummaryCompany == false)
                            {
                                BoolSummaryCompany = true;
                            }
                        }
                    }
                }
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedCompany = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisHistoryList()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var historyRegex = new Regex(@"_history[0-9]");

            // Reuse checking
            var HistoryQuery = RawDataTable.AsEnumerable()
                .Where(r => r.Field<string>("Usernames").Contains("_history"));

            // Temp Table to count reuse occurrences
            DataTable DataTableHistoryReuse_Temp = new DataTable();
            DataTableHistoryReuse_Temp.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Current Passwords", typeof(string)),
                               new DataColumn("History Username", typeof(string)),
                });

            if (HistoryQuery.Count() > 0)
            {

                foreach (DataRow row in RawDataTable.Rows)
                {
                    string user = row[0].ToString();
                    string pass = row[1].ToString();

                    if (historyRegex.Match(user).Success == false)
                    {
                        DataTableHistoryAll.Rows.Add(user, pass);

                        foreach (DataRow HistoryUser in HistoryQuery)
                        {
                            string HistoryUsername = HistoryUser.ItemArray[0].ToString();
                            string HistoryPassword = HistoryUser.ItemArray[1].ToString();

                            if (HistoryUsername.Split('_')[0] == (user) && pass == HistoryPassword) // Fix for mis-matched names in history
                            //if (HistoryUsername.Contains(user) && pass == HistoryPassword)
                            {
                                DataTableHistoryReuse_Temp.Rows.Add(user, pass, HistoryUsername);
                                count++;
                            }
                        }
                    }
                }


                // Count occurences of history reuse

                var q = DataTableHistoryReuse_Temp.Rows.OfType<DataRow>()
                    .GroupBy(x => x[0].ToString())
                    .Select(g => new { Username = g.Key, Count = g.Count(), Rows = g.ToList() })
                    .OrderByDescending(x => x.Count);

                foreach (var z in q)
                {
                    // Password Retrieval
                    var HistoryQueryPasswordTemp = RawDataTable.AsEnumerable()
                        .Where(r => r.Field<string>("Usernames").Equals(z.Username));
                    List<DataRow> list = HistoryQueryPasswordTemp.AsEnumerable().ToList();
                    string HistoryQueryPassword = list[0].ItemArray[1].ToString();
                    DataTableHistoryReuse.Rows.Add(z.Username, HistoryQueryPassword, z.Count);
                }

                if (DataTableHistoryReuse.Rows.Count > 1)
                {
                    BoolSummaryHistory = true;
                }

                DataView dv = DataTableHistoryAll.DefaultView;
                dv.Sort = "Usernames asc";
                DataTable sortedDT = dv.ToTable();
            }

            //return 
            Dispatcher.Invoke(() => DataGridHistoryReuse.ItemsSource = DataTableHistoryReuse.AsDataView());

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedHistory = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisWeakPassword()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;
            var history_regex = new Regex(@"_history[0-9]");
            var weak_regex = new Regex(@"[a-zA-Z\d!@#$%^&*()_+-=[]\s{};':|,.<>\/?]{ 10, }");

            var lower_regex = new Regex(@"[a-z]");
            var higher_regex = new Regex(@"[A-Z]");
            var numeric_regex = new Regex(@"[0-9]");
            var special_regex = new Regex(@"[!@£$%^&*()\[\]:;\\\/<>{}' ]");
            var repeat_regex = new Regex(@"([A-Z\d])\1\1");
            var sequential_regex = new Regex(@"(abc|bcd|cde|def|efg|fgh|ghi|hij|ijk|jkl|klm|lmn|mno|nop|opq|pqr|qrs|rst|stu|tuv|uvw|vwx|wxy|xyz|012|123|234|345|456|567|678|789|987|876|765|654|543|432|321|210|zyx|yxw|xwv|wvu|vut|uts|tsr|srq|rqp|qpo|pon|onm|nml|mlk|lkj|kji|jih|ihg|hgf|gfe|fed|edc|dcb|cba)");


            bool boolWeak = false;

            foreach (DataRow row in RawDataTable.Rows)
            {
                string username = row[0].ToString();
                string password = row[1].ToString();

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                }

                if (password.Length < 10)
                {
                    if ((lower_regex.Match(password).Success == false) && (higher_regex.Match(password).Success == false) && (numeric_regex.Match(password).Success == false) && (special_regex.Match(password).Success == false))
                    {
                        boolWeak = true;
                    }
                    else if (repeat_regex.Match(password).Success)
                    {
                        boolWeak = true;
                    }
                    else if (sequential_regex.Match(password).Success)
                    {
                        boolWeak = true;
                    }

                    if (boolWeak)
                    {
                        count++;
                        DataTableWeak.Rows.Add(username, password);
                        if (BoolSummaryWeak == false)
                        {
                            BoolSummaryWeak = true;
                        }
                    }

                }
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedWeak = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisHashcatMask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            var lower_regex = new Regex(@"[a-z]");
            var higher_regex = new Regex(@"[A-Z]");
            var numeric_regex = new Regex(@"[0-9]");
            var special_regex = new Regex(@"[!@£$%^&*()\[\]:;\\\/<>{}]");
            var history_regex = new Regex(@"_history[0-9]");


            DataTable DataTableHashcatTemp = new DataTable();
            DataTableHashcatTemp.Columns.AddRange(new[]
                                       {
                               new DataColumn("Hashcat Mask", typeof(string)),
                               new DataColumn("Occurences", typeof(int)),
            });

            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data
                string mask = "";
                string username = row[0].ToString();
                string password = row[1].ToString();

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                }

                for (int i = 0; i < password.Length; i++)
                {
                    string c = password[i].ToString();
                    //Process char
                    if ((lower_regex.Match(c).Success))
                    {
                        mask = mask + "?l";
                    }
                    else if ((higher_regex.Match(c).Success))
                    {
                        mask = mask + "?u";
                    }
                    else if ((numeric_regex.Match(c).Success))
                    {
                        mask = mask + "?d";
                    }
                    else if ((special_regex.Match(c).Success))
                    {
                        mask = mask + "?s";
                    }
                }

                DataTableHashcatTemp.Rows.Add(mask.ToString());
            }

            var q = DataTableHashcatTemp.Rows.OfType<DataRow>()
                .GroupBy(x => x[0].ToString())
                .Select(g => new { hashmask = g.Key, Count = g.Count(), Rows = g.ToList() })
                .OrderByDescending(x => x.Count);

            foreach (var z in q)
            {
                DataRow dataRow = DataTableHashcatMasks.NewRow();
                dataRow[0] = z.hashmask;
                dataRow[1] = z.Count;

                DataTableHashcatMasks.Rows.Add(dataRow);

            }
            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedHashcat = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public void AnalysisCharacters()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int count = 0;

            var regex_alphaCharLower = new Regex(@"[a-z]");
            var regex_alphaCharUpper = new Regex(@"[A-Z]");
            var regex_numChar = new Regex(@"[0-9]");
            var regex_specChar = new Regex(@"[^A-Za-z0-9]"); // ['!', '@', '£', '#', '$', '%', '^', '&', '*', '(', ')', '-', '+', '=', '_', '|', '?', '`', '±', '§', ';', ':']
            var history_regex = new Regex(@"_history[0-9]");

            // Break out of method if no results
            if (RawDataTable.Rows.Count == 0)
            {
                Dispatcher.Invoke(() => LabelStatusCharacter.Text = ("0 Records"));
                return;
            }

            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data
                string username = row[0].ToString();
                string password = row[1].ToString();

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                    //return;
                }

                foreach (char passChar in password)
                {
                    if (regex_alphaCharLower.Match(passChar.ToString()).Success)
                    {
                        DataTableCharacterFrequency.Rows.Add(passChar, "Lower Alpha");
                    }
                    else if (regex_alphaCharUpper.Match(passChar.ToString()).Success)
                    {
                        DataTableCharacterFrequency.Rows.Add(passChar, "Upper Alpha");
                    }
                    else if (regex_numChar.Match(passChar.ToString()).Success)
                    {
                        DataTableCharacterFrequency.Rows.Add(passChar, "Numeric");
                    }
                    else if (regex_specChar.Match(passChar.ToString()).Success)
                    {
                        DataTableCharacterFrequency.Rows.Add(passChar, "Special");
                    }
                };
            };

            // Individual Character Sort
            var res = (from x in DataTableCharacterFrequency.AsEnumerable()
                       group x by (string)x["Characters"] into y
                       select new { y.Key, Count = y.Count() }).ToArray();

            // Type Sort
            var type = (from x in DataTableCharacterFrequency.AsEnumerable()
                        group x by (string)x["Character Type"] into y
                        select new { y.Key, Count = y.Count() }).ToArray();

            // Character type frequency
            foreach (var z in res)
            {
                DataTableCharacterTemp.Rows.Add(z.Key, z.Count);

                count++;
                if (BoolSummaryCharacter == false)
                {
                    BoolSummaryCharacter = true;
                }
            }

            foreach (var z in type)
            {
                DataTableCharacterTypeTemp.Rows.Add(z.Key, z.Count);
            }


            // Character Type Output and Count

            DataView dvCharType = DataTableCharacterTypeTemp.DefaultView;
            dvCharType.Sort = "Occurrences desc";
            DataTable sortedDTCharType = dvCharType.ToTable();

            StringCharacterType = sortedDTCharType.Rows[0][0].ToString();
            DataTableCharacterType = sortedDTCharType;

            // Character Occurrences Output

            DataView dvChar = DataTableCharacterTemp.DefaultView;
            dvChar.Sort = "Occurrences desc";
            DataTable sortedDTChar = dvChar.ToTable();

            DataTableCharacter = sortedDTChar;

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            TimeElapsedCharacter = elapsedTime.ToString();
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        public async Task AnalysisPasswordSuffix()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            var regex_alphaCharLower = new Regex(@"[a-z]");
            var regex_alphaCharUpper = new Regex(@"[A-Z]");
            var regex_numChar = new Regex(@"[0-9]");
            var regex_specChar = new Regex(@"[^A-Za-z0-9]"); // ['!', '@', '£', '#', '$', '%', '^', '&', '*', '(', ')', '-', '+', '=', '_', '|', '?', '`', '±', '§', ';', ':']
            var history_regex = new Regex(@"_history[0-9]");

            // Break out of method if no results
            if (RawDataTable.Rows.Count == 0)
            {
                return;
            }

            foreach (DataRow row in RawDataTable.Rows)
            {
                // Set from raw data
                string username = row[0].ToString();
                string password = row[1].ToString();

                if (BoolIncludeHistory == false && history_regex.Match(username).Success)
                {
                    continue;
                    //return;
                }

                string reverseCharString = "";

                // Skips password if does not contain alpha chars
                if (!regex_alphaCharUpper.Match(password).Success && !regex_alphaCharLower.Match(password).Success)
                {
                    continue;
                }

                foreach (char passChar in password.Reverse())
                {
                    if (regex_alphaCharLower.Match(passChar.ToString()).Success)
                    {
                        break;
                    }
                    else if (regex_alphaCharUpper.Match(passChar.ToString()).Success)
                    {
                        break;
                    }
                    else if (regex_numChar.Match(passChar.ToString()).Success)
                    {
                        reverseCharString += passChar.ToString();
                    }
                    else if (regex_specChar.Match(passChar.ToString()).Success)
                    {
                        reverseCharString += passChar.ToString();
                    }
                }
                if (reverseCharString.Length > 0)
                {
                    string SuffixString = ReverseString(reverseCharString).ToString();
                    DataTableTrailingSuffixTemp.Rows.Add(SuffixString);
                }
            };

            var res = (from x in DataTableTrailingSuffixTemp.AsEnumerable()
                       group x by (string)x["Suffix"] into y
                       select new { y.Key, Count = y.Count() }).ToArray();

            foreach (var z in res)
            {
                DataTableTrailingSuffix.Rows.Add(z.Key, z.Count);
                DataTableTrailingSuffixMask.Rows.Add(HashcatSuffixRule(z.Key), z.Count);
            }

            DataView dvSuffix = DataTableTrailingSuffix.DefaultView;
            dvSuffix.Sort = "Occurrences desc";
            DataTable sortedDTTrailingSuffix = dvSuffix.ToTable();

            DataView dvSuffixMask = DataTableTrailingSuffixMask.DefaultView;
            dvSuffixMask.Sort = "Occurrences desc";
            DataTable sortedDTRuleSuffix = dvSuffixMask.ToTable();

            DataTableTrailingSuffixMask = sortedDTRuleSuffix;
            DataTableTrailingSuffix = sortedDTTrailingSuffix;
            ProgressBarUpdate();
            SetTimeElapsedBinding();
        }

        // End Analysis Section

        async Task PutTaskDelay()
        {
            await Task.Delay(200);
        }

        // Summary Outputs

        private void GenerateSummary()
        {
            ButtonGenerateSummary.Background = Brushes.DimGray;
            ButtonGenerateSummary.Content = "Please Wait";
            ButtonGenerateSummary.IsEnabled = false;

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            if (CheckSummaryExec.IsChecked.Value)
            {
                Dispatcher.Invoke(() => SummaryOutputExecutive());

            }
            if (CheckSummaryTech.IsChecked.Value)
            {
                Dispatcher.Invoke(() => RichTextTech.Text = "");
                Dispatcher.Invoke(() => RichTextXML.Text = "");
                Dispatcher.Invoke(() => SummaryOutputTechnical());
            }
            if (CheckSummaryRecommend.IsChecked.Value)
            {
                Dispatcher.Invoke(() => RichTextRecommend.Text = "");
                Dispatcher.Invoke(() => SummaryOutputRecommend());
            }

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

            BrushConverter bc = new BrushConverter();
            Brush brush = (Brush)bc.ConvertFrom("#FF03A9F4");
            brush.Freeze();
            ButtonGenerateSummary.Background = brush;
            ButtonGenerateSummary.IsEnabled = true;
            ButtonGenerateSummary.Content = "Generate Summary";
        }

        private void SummaryOutputExecutive()
        {

            float totalPasswords;
            RichTextExec.Text = "";
            if (BoolIncludeHistory)
            {
                totalPasswords = float.Parse(RawDataTable.Rows.Count.ToString());
            }
            else
            {
                totalPasswords = RawDataTable.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
            }
            string ExecEnabledString = "";

            // Intro
            if (ADUsers.Rows.Count > totalPasswords)
            {
                int AdUsersEnabledOrAllCount = 0;
                string ExecUsersTotalPercent = "";
                if (RadioAll.IsChecked == true)
                {
                    AdUsersEnabledOrAllCount = ADUsers.Rows.Count;
                    float AllPercentRaw = (totalPasswords / AdUsersEnabledOrAllCount);
                    ExecUsersTotalPercent = Math.Round(AllPercentRaw * 100, 0).ToString();
                    ExecEnabledString = "This amounts to " + ExecUsersTotalPercent + "% of the " + AdUsersEnabledOrAllCount.ToString() + " total domain user accounts that were able to be computed into plaintext.";
                }
                else
                {
                    AdUsersEnabledOrAllCount = ADUsers.AsEnumerable().Where(c => (c["Account Status"].ToString().Contains("Enabled"))).ToList().Count;
                    float AllPercentRaw = (totalPasswords / AdUsersEnabledOrAllCount);
                    ExecUsersTotalPercent = Math.Round(AllPercentRaw * 100, 0).ToString();
                    ExecEnabledString = "This amounts to " + ExecUsersTotalPercent + "% of the " + AdUsersEnabledOrAllCount.ToString() + " total enabled domain user accounts that were able to be computed into plaintext.";
                }
            }
            ExecutiveSummaryContent = ("A password audit was performed against the extracted password hashes. Password cracking methods and tools were used to enumerate the plaintext passwords, and as such not all of the passwords were able to be identified within a reasonable time frame. In total, there were " + (totalPasswords) + " username and plaintext password combinations that were obtained and have been analysed. " + ExecEnabledString);

            ExecutiveSummaryContent += "\n";
            ExecutiveSummaryContent += "\n";
            bool listSummaryCommon = false;

            var FreqRows = GetDataGridRows(DataGridFrequency);

            // Password Frequency
            int count = 0;
            foreach (DataRowView rowView in DataGridFrequency.Items)
            {
                string sumString;
                if (count == 0)
                {
                    if (int.Parse(rowView.Row[1].ToString()) > 1)
                    {
                        ExecutiveSummaryContent += ("As part of the password audit, the top 10 most commonly used passwords within the organisation have been compiled. This list has been broken up with the most frequently used passwords and the numeric value of the total passwords for each: ");
                        listSummaryCommon = true;
                    }
                    else
                    {
                        ExecutiveSummaryContent += ("As part of the password audit, the top 10 most commonly used passwords within the organisation have been compiled. Throughout the list of passwords that were analysed there were no multiple occurrences of passwords, highlighting that each user account has a unique password. This does not mean that uncracked passwords also share unique passwords, and as such an analysis of the hashes should also be reviewed.");
                        listSummaryCommon = false;
                    }
                }

                if (count < 10 && listSummaryCommon)
                {
                    float freqCount = float.Parse(rowView.Row[1].ToString());
                    float freqPercentRaw = (freqCount / totalPasswords);
                    string freqPercentage = Math.Round(freqPercentRaw * 100, 0).ToString();
                    if (freqPercentage == "0")
                    {

                        freqPercentage = "< 1";

                    }
                    string topPass = rowView.Row[0].ToString();
                    if (topPass.Length == 0)
                    {
                        topPass = "*****BLANK PASSWORD*****";
                    }
                    sumString = ("- " + topPass + " : " + rowView.Row[1].ToString() + " : " + freqPercentage + "%");
                    ExecutiveSummaryContent += "\n";
                    ExecutiveSummaryContent += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            // Company Name
            if (BoolSummaryCompany && ListBoxCompany.Items.Count > 0)
            {
                string multiOrgName = "";
                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += "\n";
                if (ListBoxCompany.Items.Count > 1)
                {
                    foreach (string orgname in ListBoxCompany.Items)
                    {

                        multiOrgName = FirstLetterToUpper(orgname) + ", " + multiOrgName;
                    }
                }
                else
                {
                    foreach (string orgname in ListBoxCompany.Items)
                    {

                        multiOrgName = FirstLetterToUpper(orgname);
                    }
                }

                ExecutiveSummaryContent += ("The organisation name, or a variation of the name (such as an abbreviation) '" + multiOrgName + "' was found to appear within " + (DataTableCompany.Rows.Count.ToString()) + " of the passwords that were able to be obtained during the password audit. For any system or administrative user accounts that have a variation of the company name as their password it is highly recommended that the passwords are changed to prevent targeted guessing attacks.");
            }

            // Good commentary section
            else if (ListBoxCompany.Items.Count > 0 && BoolSummaryCompany == false)
            {
                string multiOrgName = "";
                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += "\n";

                if (ListBoxCompany.Items.Count > 1)
                {
                    foreach (string orgname in ListBoxCompany.Items)
                    {

                        multiOrgName = FirstLetterToUpper(orgname) + ", " + multiOrgName;
                    }
                }
                else
                {
                    foreach (string orgname in ListBoxCompany.Items)
                    {

                        multiOrgName = FirstLetterToUpper(orgname);
                    }
                }
                ExecutiveSummaryContent += ("The organisation name, or a variation of the name (such as an abbreviation) '" + multiOrgName + "' was not found to appear within any of the passwords that were able to be obtained during the password audit. This demonstrates that staff members and other users have been educated sufficiently to the risks of using a variation of the organisation name within their passwords.");
            }
            ExecutiveSummaryContent += "\n";


            // Common base Words
            if (BoolSummaryCommon)
            {
                ExecutiveSummaryContent += "\n";
                int CommonCount = DataGridCommon.Items.Count;
                string WordChoice;

                if (CommonCount == 1)
                {
                    WordChoice = "was ";
                }
                else
                {
                    WordChoice = "were ";
                }

                ExecutiveSummaryContent += ("One of the most common weaknesses relating to passwords is where users and administrators utilise passwords that are a variation of commonly used passwords, base dictionary words, names, and even phrases. Overall, there " + WordChoice + (DataGridCommon.Items.Count.ToString()) + " passwords that were found to have a variation of one of these common words or phrases. Some of these passwords are based on the words 'password', 'qwerty', 'starwars', 'system', 'admin', 'letmein', and 'iloveyou'.");
                ExecutiveSummaryContent += "\n";
            }

            int KeyboardCount = DataTableKeyboard.Rows.Count;
            // Keyboard Walking
            if (BoolSummaryKeyboard)
            {
                string WordChoice;


                if (KeyboardCount == 1)
                {
                    WordChoice = "was ";
                }
                else
                {
                    WordChoice = "were ";
                }
                ExecutiveSummaryContent += "\n";

                ExecutiveSummaryContent += ("As part of the wider password analysis each password was assessed and compared to common keyboard patterns. These keyboard patterns were defined by the QWERTY, AZERTY, or DVORAK layouts and are noted where a password is made up of characters in close proximity such as 'qwer', 'zxcvbn', and 'qazwsx' as an example. This is known as 'keyboard walking'. In total there " + WordChoice + (KeyboardCount.ToString()) + " passwords in use that had at least one of these or other variations.");
                ExecutiveSummaryContent += "\n";
            }
            else
            {
                ExecutiveSummaryContent += "\n";

                ExecutiveSummaryContent += ("As part of the wider password analysis each user account password was assessed and compared to common keyboard patterns. These keyboard patterns were defined by the QWERTY, AZERTY, or DVORAK layouts and are noted to be where a password is made up of characters in close proximity such as where users have used 'qwer', 'zxcvbn', and 'qazwsx' as an example. This is known as 'keyboard walking'. The results of this part of the audit revealed that there were no verifiable instances of user accounts being set with a common keyboard walking pattern. ");
                ExecutiveSummaryContent += "\n";
            }

            // Username in Password
            if (BoolSummaryUser)
            {
                string WordChoice;
                int UserPassword = DataTableUsernames.Rows.Count;

                if (UserPassword == 1)
                {
                    WordChoice = "was ";
                }
                else
                {
                    WordChoice = "were ";
                }
                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += ("There " + WordChoice + (DataGridUsernames.Items.Count.ToString()) + " passwords that were identified as using a password that was a variation of the username. This includes additional prefixed or suffixed characters, substitutions within the word (i.e. 3 instead of e), or the username as it appears. Penetration testers, and more importantly attackers, will often check system or administrative accounts that have a variation of the username set as the password and as such it is critical that organisations do not use this convention for password security.");
                ExecutiveSummaryContent += "\n";
            }
            else
            {
                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += ("There were no instances that were identified where user accounts were using a password that was a variation of the username. This area of analysis focuses on passwords that have additional prefixed or suffixed characters, substitutions within the word (i.e. 3 instead of e), but also includes verifying whether the username has directly been set as the password. Penetration testers, and more importantly attackers will often check system or administrative accounts that have a variation of the username set as the password, and as such it is critical that organisations do not use this convention for password security.");
                ExecutiveSummaryContent += "\n";
            }

            // Password Reuse/sharing
            if (BoolSummaryReuse)
            {
                var ReuseRows = GetDataGridRows(DataGridReuse);
                int ReuseRowCount = DataGridReuse.Items.Count;
                string WordChoice;
                string WordAccount;
                if (ReuseRowCount == 1)
                {
                    WordChoice = "was ";
                    WordAccount = "account";
                }
                else
                {
                    WordChoice = "were ";
                    WordAccount = "accounts";
                }
                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += ("There " + WordChoice + ReuseRowCount.ToString() + " user " + WordAccount + " that were found to share the same password between similarly named accounts. These are believed to be service accounts or individual user accounts that are operated by the same user. Password re-use should be investigated as when a password is reused between a low privileged and a high privileged account it can lead to a compromise of systems that the high privileged account is authorised to access.");
                ExecutiveSummaryContent += "\n";

                if (DataTableReuseAdmin.Rows.Count > 0)
                {
                    var AdminReuseRows = GetDataGridRows(DataGridReuseAdmin);
                    int AdminReuseRowCount = DataGridReuseAdmin.Items.Count;
                    string AdminWordChoice;
                    string AdminWordAccount;
                    if (AdminReuseRowCount == 1)
                    {
                        AdminWordChoice = "was ";
                        AdminWordAccount = "account";
                    }
                    else
                    {
                        AdminWordChoice = "were ";
                        AdminWordAccount = "accounts";
                    }
                    ExecutiveSummaryContent += "\n";
                    ExecutiveSummaryContent += ("During the audit there " + AdminWordChoice + AdminReuseRowCount.ToString() + " user " + AdminWordAccount + " that were identified as sharing the same password as other administrative accounts. This can often be an indication that service or administrative accounts are not being set with strong and unique passwords.");
                    ExecutiveSummaryContent += "\n";
                }
            }

            // Dates
            if (BoolSummaryDate)
            {
                ExecutiveSummaryContent += "\n";
                string DateRowsCount = DataGridDateTime.Items.Count.ToString();
                ExecutiveSummaryContent += ("Users often utilise dates and times when creating their passwords. This usually appears in a basic form, such as 'Tuesday123', 'Summer2018', 'January18', etc. As part of this audit it was possible to discover " + DateRowsCount + " passwords using a variation of these basic dates/times. Depending upon the simplicity of the passwords it may be trivial for an attacker to brute force these and gain access to enabled user accounts.");
                ExecutiveSummaryContent += "\n";
            }
            else
            {
                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += ("Users often utilise dates and times when creating their passwords. This usually appears in a basic form, such as 'Tuesday123', 'Summer2018', 'January18', etc. There were no instances where user accounts used passwords that contained a form of the date (e.g. day of the week, month, season, etc).");
                ExecutiveSummaryContent += "\n";
            }


            /* History Repeat Password */

            //string sumString;
            if (BoolSummaryHistory)
            {
                float historyCount = float.Parse(DataTableHistoryReuse.Rows.Count.ToString());
                float historyPercentRaw = (historyCount / totalPasswords);
                string historyPercentage = Math.Round(historyPercentRaw * 100, 0).ToString();

                string WordChoice;
                string WordAccount;
                if (historyCount == 1)
                {
                    WordChoice = "was ";
                    WordAccount = "account";
                }
                else
                {
                    WordChoice = "were ";
                    WordAccount = "accounts";
                }

                if ((int.Parse(historyCount.ToString())) > 1)
                {
                    ExecutiveSummaryContent += "\n";
                    ExecutiveSummaryContent += ("This audit included a review of historical Active Directory passwords. These user accounts are suffixed with '_history' and are numbered incrementally from the most recent password to the oldest stored in Active Directory (for example user1, user1_history0, user1_history1, etc). In total there " + WordChoice + historyCount.ToString() + " user " + WordAccount + " that " + WordChoice + "found to re-use a previous password.");
                    ExecutiveSummaryContent += "\n";
                }
            }

            // Character Anaylsis Entropy
            if (BoolSummaryCharacter)
            {
                string CharacterEntropyExecText = "";

                int CharacterCount = DataGridCharacter.Items.Count;
                if (CharacterCount >= 85)
                {
                    CharacterEntropyExecText = "This character-set was found to be largely varied, likely indicating that users utilise uppercase, lowercase, numeric, and special characters. The larger the character variance the potentially more random and complex the user password.";
                }
                else if (CharacterCount < 85 && CharacterCount > 36)
                {
                    CharacterEntropyExecText = "The passwords that were identified were found to be using a mixture of uppercase, lowercase, numeric, and special characters although there is room for further improvement in the variety of characters that are used. Users should be educated to include different special characters within their passwords, which would considerably increase the total entropy and reduce the likelihood of their passwords being compromised in the event of a malicious attack or data breach.";
                }
                else if (CharacterCount <= 36)
                {
                    CharacterEntropyExecText = "Users do not appear to be using a large variance of available characters within their passwords, which greatly reduces the potential entropy and increases the likelihood than a malicious attacker could compromise a password in the event of a brute-force or offline-based cracking attack. Depending upon the total percentage of passwords that were able to be obtained for the analysis aspect this could indicate that poor password practices are being followed throughout the organisation.";
                }

                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += ("Following character analysis of the entire password list, which was comprised of a total number of " + (CharacterCount.ToString()) + " different characters from various character-sets, the majority of users were mainly found to utilise " + StringCharacterType + " characters. " + CharacterEntropyExecText);
                ExecutiveSummaryContent += "\n";
            }

            // Pwned Password Summary
            if (BoolSummaryPwned)
            {
                int CharacterCount = DataGridPwned.Items.Count;

                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += ("Password cracking attacks are most often catalysed by the use of 'wordlists', also known as a dictionary attack, wherein a list of words, phrases, or more commonly previously discovered passwords are used to compare and compute the password hashes to a plaintext counterpart. Further analysis performed against the top 10 most frequently used passwords as a result of this wider password analysis showed that there were a number of passwords that were known to have been compromised in publicly released data breaches.");
                ExecutiveSummaryContent += "\n";
            }

            // Pwned Admin Password Summary
            if (BoolSummaryPwned && DataTablePwnedAdmin.Rows.Count > 0)
            {
                int CharacterCount = DataTablePwnedAdmin.Rows.Count;
                string word = "was";
                if (CharacterCount > 1)
                {
                    word = "were";

                }

                ExecutiveSummaryContent += "\n";
                ExecutiveSummaryContent += ("Additionally, there " + word + " " + CharacterCount.ToString() + " administrative user accounts that were found to be using passwords that have previously been seen in data leaks. The risk of administrative accounts using these passwords is that it would be trivial for an attacker to compromise the respective accounts using publicly available wordlists that have been created as a result of these data breaches.");
                ExecutiveSummaryContent += "\n";
            }

            // Administrators / Services
            if (BoolSummaryAdmin)
            {
                ExecutiveSummaryContent += "\n";
                string adminWordChoice;
                string adminWordAccount;
                var AdminRows = GetDataGridRows(DataGridAdmin);

                int adminRowCount = DataGridAdmin.Items.Count;
                if (adminRowCount == 1)
                {
                    adminWordChoice = "was ";
                    adminWordAccount = "account";
                }
                else
                {
                    adminWordChoice = "were ";
                    adminWordAccount = "accounts";
                }
                ExecutiveSummaryContent += ("Finally, there " + adminWordChoice + adminRowCount.ToString() + " administrative " + adminWordAccount + " (built-in Administrators, Domain Admins, Enterprise Admins, etc.) that " + adminWordChoice + "able to be compromised through password cracking. Administrative user accounts (including service accounts with administrative permissions) should use unique, strong passwords. Using a password manager for these accounts would ensure that long, random passwords can be easily used and would therefore greatly increase the security posture of these user accounts.");

            }

            Dispatcher.Invoke(() => RichTextExec.Text = (ExecutiveSummaryContent));
        }

        private void SummaryOutputTechnical()
        {
            TechnicalSummaryContent = "";
            TechnicalSummaryContentMD = "";
            bool TechMarkdownOutput = false;
            if (CheckSummaryMarkdown.IsChecked == true)
            {
                TechMarkdownOutput = true;
            }
            ConfirmXMLOutputSettingsMissing();

            string StringSummaryXMLTemp;
            // Temp create XML Format
            StringSummaryXMLOutput = " <?xml version='1.0' encoding='utf8'?>";
            StringSummaryXMLOutput += "\n<items source=\"" + Properties.Settings.Default.VulnIdSource + "\" version=\"1.0\">";
            StringSummaryXMLOutput += "\n\t<item ipaddress=\"\" hostname=\"" + SummaryXMLDomain.Text + "\">";
            StringSummaryXMLOutput += "\n\t\t<services>";
            StringSummaryXMLOutput += "\n\t\t\t<service name=\"\" port=\"\" protocol=\"tcp\">";
            StringSummaryXMLOutput += "\n\t\t\t\t<vulnerabilities>";

            float totalPasswords;



            if (BoolIncludeHistory)
            {
                totalPasswords = float.Parse(RawDataTable.Rows.Count.ToString());
            }
            else
            {
                totalPasswords = RawDataTable.AsEnumerable().Where(c => (!c["Usernames"].ToString().Contains("_history"))).ToList().Count;
            }

            StringSummaryXMLTemp = "A password audit was performed against the extracted password hashes. Password cracking methods and tools were used to enumerate the plaintext password counterparts, and as such not all of the passwords were able to be identified. In total, there were " + (totalPasswords) + " username and password combinations that were obtained and have been analysed.";
            StringSummaryXMLTemp += "\n";
            StringSummaryXMLTemp += "\n";
            StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdOverview, EncodeXMLChars(StringSummaryXMLTemp));
            TechnicalSummaryContent += StringSummaryXMLTemp;


            bool listSummaryCommon = false;
            int count = 0;

            /* Most Frequent Characters */
            foreach (DataRow dr in DataTableCharacter.Rows)
            {
                string sumString;


                if (count == 0)
                {
                    if (int.Parse(dr.ItemArray[1].ToString()) > 0)
                    {
                        StringSummaryXMLTemp = "The top 20 characters used out of " + totalPasswords + " passwords and the total number of instances (character and total number of instances): ";
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n| Characters | Times Seen |\n| -------- | -------- |";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (count < 20 && listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + dr.ItemArray[0].ToString() + " |" + dr.ItemArray[1].ToString() + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + dr.ItemArray[1].ToString());
                    }

                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n\n";
            }

            bool listSummaryCommonType = false;
            count = 0;

            foreach (DataRow dr in DataTableCharacterType.Rows)
            {
                string sumString;
                if (count == 0)
                {
                    StringSummaryXMLTemp += ("From the " + totalPasswords + " passwords that were assessed the overall usage of each different character-type used within passwords has been assessed and is broken down as follows: ");
                    if (TechMarkdownOutput)
                    {
                        StringSummaryXMLTemp += "\n\n| Character Type | Times Seen |\n| -------- | -------- |";
                    }
                    listSummaryCommonType = true;
                }

                if (listSummaryCommonType)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + dr.ItemArray[0].ToString() + " | " + dr.ItemArray[1].ToString() + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + dr.ItemArray[1].ToString());
                    }

                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }
            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
            }


            listSummaryCommon = false;
            count = 0;

            /* Most Frequent Passwords */
            foreach (DataRow dr in DataTableFrequency.Rows)
            {
                string sumString;


                if (count == 0)
                {
                    if (int.Parse(dr.ItemArray[1].ToString()) > 1)
                    {
                        StringSummaryXMLTemp += ("As part of the password audit, the top 15 most commonly used passwords within the organisation has been compiled. This list has been broken up with the most frequently used passwords and the numeric value of the total passwords for each: ");
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n| Password | Times Seen | Percentage of Total |\n| --------| -------- | -------- |";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (count < 15 && listSummaryCommon)
                {

                    float freqCount = float.Parse(dr.ItemArray[1].ToString());
                    float freqPercentRaw = (freqCount / totalPasswords);
                    string freqPercentage = Math.Round(freqPercentRaw * 100, 0).ToString();
                    if (freqPercentage == "0")
                    {

                        freqPercentage = "< 1";

                    }
                    string topPass = dr.ItemArray[0].ToString();
                    if (topPass.Length == 0)
                    {
                        topPass = "*****BLANK PASSWORD*****";
                    }
                    if (TechMarkdownOutput)
                    {
                        sumString = ("| " + topPass + " | " + dr.ItemArray[1].ToString() + " | " + freqPercentage + "% |");
                    }
                    else
                    {
                        sumString = ("- " + topPass + " : " + dr.ItemArray[1].ToString() + " : " + freqPercentage + "%");
                    }

                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdCharFreq, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }


            count = 0;
            listSummaryCommon = false;

            /* Most Frequent Lengths */
            string tempMin = Properties.Settings.Default.LengthMinimumUserSetting;

            foreach (DataRow dr in DataTableLengthFreq.Rows)
            {
                string sumString;


                if (count == 0)
                {
                    if (int.Parse(dr.ItemArray[1].ToString()) > 0)
                    {
                        StringSummaryXMLTemp = ("The following is a top 10 descending list of the most popular password lengths and their representative amount of the total passwords that were analysed (total count and percentage): ");
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n| Password Length | Times Seen | Percentage of Total |\n| --------| -------- | -------- |";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (count < 10 && listSummaryCommon)
                {
                    float freqCount = float.Parse(dr.ItemArray[1].ToString());
                    float freqPercentRaw = (freqCount / totalPasswords);
                    string freqPercentage = Math.Round(freqPercentRaw * 100, 0).ToString();
                    if (freqPercentage == "0")
                    {

                        freqPercentage = "< 1";

                    }
                    if (TechMarkdownOutput)
                    {
                        sumString = ("| " + (dr.ItemArray[0].ToString()) + " | " + dr.ItemArray[1].ToString() + " | " + freqPercentage + "%|");
                    }
                    else
                    {
                        sumString = ("- Length : " + (dr.ItemArray[0].ToString()) + " : " + dr.ItemArray[1].ToString() + " : " + freqPercentage + "%");
                    }

                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                if (DataTableLength.Rows.Count == 0)
                {
                    StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdLength, EncodeXMLChars(StringSummaryXMLTemp));
                    TechnicalSummaryContent += StringSummaryXMLTemp;

                }
            }
            count = 0;
            listSummaryCommon = false;


            /* Length Non-Compliance */
            foreach (DataRow dr in DataTableLength.Rows)
            {
                string sumString;


                if (count == 0)
                {
                    if ((dr.ItemArray[1].ToString().Length) > 0)
                    {
                        StringSummaryXMLTemp += ("The length of the following user accounts have passwords set that do not meet the minimum of " + tempMin.ToString() + " characters: ");
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n| Username | Password |\n| -------- | -------- |";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                    }

                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(SummaryXMLVulnLength.Text, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }
            listSummaryCommon = false;
            count = 0;



            /* Username as Password */
            foreach (DataRow dr in DataTableUsernames.Rows)
            {
                string sumString;


                if (count == 0)
                {
                    if ((dr.ItemArray[1].ToString().Length) > 0)
                    {
                        StringSummaryXMLTemp = ("The following user accounts were found to have the username, or a variation of the username, set as the password: ");
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n| Username | Password | \n| -------- | -------- |";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + dr.ItemArray[0].ToString() + " |" + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                    }

                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdUsername, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }
            listSummaryCommon = false;
            count = 0;

            string StringSumamryXMLReuse = "";
            /* Shared Passwords - Similar Account Names*/
            foreach (DataRow dr in DataTableReuse.Rows)
            {
                string sumString;


                if (count == 0)
                {
                    if ((dr.ItemArray[1].ToString().Length) > 0)
                    {
                        StringSumamryXMLReuse += ("The following accounts were found to share the same password between similarly named accounts. These are believed to be service accounts or individual user accounts that are operated by the same user. Password re-use should be investigated as when a password is reused between a low privileged and a high privileged account it can lead to a compromise of systems that the high privileged account is authorised to access: ");
                        if (TechMarkdownOutput)
                        {
                            StringSumamryXMLReuse += "\n\n| Username | Password | Shared with Username|\n| -------- | -------- | -------- |";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + " | " + dr.ItemArray[2].ToString() + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()) + " - Shared with: " + dr.ItemArray[2].ToString());
                    }
                    StringSumamryXMLReuse += "\n";
                    StringSumamryXMLReuse += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSumamryXMLReuse += "\n";
                StringSumamryXMLReuse += "\n";
            }

            if (TechMarkdownOutput)
            {
                StringSummaryXMLTemp += StringSumamryXMLReuse;
            }

            listSummaryCommon = false;
            count = 0;

            /* Shared Passwords - Admins*/
            foreach (DataRow dr in DataTableReuseAdmin.Rows)
            {
                string sumString;


                if (count == 0)
                {
                    if ((dr.ItemArray[1].ToString().Length) > 0)
                    {
                        StringSumamryXMLReuse += ("The following accounts were found to share the same password between low privileged and high-privileged administrative accounts: ");
                        if (TechMarkdownOutput)
                        {
                            StringSumamryXMLReuse += "\n\n| Username | Password | Shared with Username|\n| -------- | -------- | -------- |";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + " | " + dr.ItemArray[2].ToString() + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()) + " - Shared with: " + dr.ItemArray[2].ToString());
                    }
                    StringSumamryXMLReuse += "\n";
                    StringSumamryXMLReuse += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSumamryXMLReuse += "\n";
                StringSumamryXMLReuse += "\n";

            }

            if (StringSumamryXMLReuse.Length > 0)
            {
                if (TechMarkdownOutput)
                {
                    StringSummaryXMLTemp += StringSumamryXMLReuse;
                }
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdReuse, EncodeXMLChars(StringSumamryXMLReuse));
                TechnicalSummaryContent += StringSumamryXMLReuse;
            }


            listSummaryCommon = false;
            count = 0;

            /* Admin Accounts */
            if (DataTableAdmin.Rows.Count > 0)
            {
                foreach (DataRow dr in DataTableAdmin.Rows)
                {
                    string sumString;

                    if ((dr.ItemArray[1].ToString().Length) > 0 && listSummaryCommon == false)
                    {
                        StringSummaryXMLTemp = ("The following user accounts were identified as service accounts or user accounts with Domain Administrator privileges (Domain Admins, Enterprise Admins, built-in Administrators, etc.) and were found to use passwords that were not sufficiently strong enough for the delegated account privileges: ");
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n| Username | Password | \n| -------- | -------- | ";
                        }
                        listSummaryCommon = true;
                    }

                    if (listSummaryCommon) // 
                    {
                        if (TechMarkdownOutput)
                        {
                            sumString = ("|" + dr.ItemArray[0].ToString() + " |" + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                        }
                        else
                        {
                            sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                        }
                        StringSummaryXMLTemp += "\n";
                        StringSummaryXMLTemp += (sumString);
                        count++;
                    }
                    else
                    {
                        break;
                    }

                }
            }

            if (listSummaryCommon)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdAdmins, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }
            listSummaryCommon = false;
            count = 0;

            /* Variation of Org Name */
            if (DataTableCompany.Rows.Count > 0)
            {
                string multiOrgName = "";
                if (ListBoxCompany.Items.Count > 1)
                {
                    foreach (string orgname in ListBoxCompany.Items)
                    {

                        multiOrgName = FirstLetterToUpper(orgname) + ", " + multiOrgName;
                    }
                }
                else
                {
                    foreach (string orgname in ListBoxCompany.Items)
                    {

                        multiOrgName = FirstLetterToUpper(orgname);
                    }
                }
                foreach (DataRow dr in DataTableCompany.Rows)
                {
                    string sumString;

                    if ((dr.ItemArray[1].ToString().Length) > 0 && listSummaryCommon == false)
                    {
                        StringSummaryXMLTemp = ("The organisation name '" + multiOrgName + "', or a variation (such as '" + FirstLetterToUpper(ListBoxCompany.Items[0].ToString()) + "2017!') was found within the following " + DataGridCompany.Items.Count.ToString() + " passwords: ");
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n| Username | Password | \n| -------- | -------- | ";
                        }
                        listSummaryCommon = true;
                    }

                    if (listSummaryCommon)
                    {
                        if (TechMarkdownOutput)
                        {
                            sumString = ("| " + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                        }
                        else
                        {
                            sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                        }
                        StringSummaryXMLTemp += "\n";
                        StringSummaryXMLTemp += (sumString);
                        count++;
                    }
                    else
                    {
                        break;
                    }

                }
            }

            if (listSummaryCommon)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdCompany, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }

            listSummaryCommon = false;
            count = 0;

            /* Common Password Use */
            foreach (DataRow dr in DataTableCommon.Rows)
            {
                string sumString;

                if ((dr.ItemArray[1].ToString().Length) > 0 && listSummaryCommon == false)
                {
                    StringSummaryXMLTemp = ("The following user accounts were found to have a password that was a variation of the most common user passwords, which can include 'password', 'letmein', '123456', 'admin', 'iloveyou', 'friday', or 'qwerty': ");
                    if (TechMarkdownOutput)
                    {
                        StringSummaryXMLTemp += "\n\n| Username | Password | \n| -------- | -------- | ";
                    }
                    listSummaryCommon = true;
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("| " + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                    }
                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdDict, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }
            listSummaryCommon = false;
            count = 0;

            /* Variation of Date/Time */
            foreach (DataRow dr in DataTableDateTime.Rows)
            {
                string sumString;

                if ((dr.ItemArray[1].ToString().Length) > 0 && listSummaryCommon == false)
                {
                    StringSummaryXMLTemp = ("The following user accounts were found to have a password that was a variation of a day or date (e.g. Monday01, September2016, or Janvier08!): ");
                    if (TechMarkdownOutput)
                    {
                        StringSummaryXMLTemp += "\n\n| Username | Password | \n| -------- | -------- | ";
                    }
                    listSummaryCommon = true;
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("| " + dr.ItemArray[0].ToString() + " |" + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                    }
                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdDate, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }

            listSummaryCommon = false;
            count = 0;

            /* Keyboard Patterns */
            foreach (DataRow dr in DataTableKeyboard.Rows)
            {
                string sumString;

                if ((dr.ItemArray[1].ToString().Length) > 0 && listSummaryCommon == false)
                {
                    StringSummaryXMLTemp = ("The following user accounts were identified as having passwords that utilise common keyboard walking patterns such as qwer, zxcvbn, qazwsx, asdf, q1w2e3, a1z2e3r4t5y6, etc: ");
                    if (TechMarkdownOutput)
                    {
                        StringSummaryXMLTemp += "\n\n| Username | Password | \n| -------- | -------- | ";
                    }
                    listSummaryCommon = true;
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("| " + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                    }
                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdKeyboard, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }

            listSummaryCommon = false;
            count = 0;

            /* Pwned Passwords */
            foreach (DataRow dr in DataTablePwned.Rows)
            {
                string sumString;

                if ((dr.ItemArray[1].ToString().Length) > 0 && listSummaryCommon == false)
                {
                    StringSummaryXMLTemp = ("Each of the following passwords, which were extracted from the most frequently used passwords in the analysis results, were identified as being compromised in previous data breaches that have been made public. The number of times that these have been seen in a data leak is listed next to the password: ");
                    if (TechMarkdownOutput)
                    {
                        StringSummaryXMLTemp += "\n\n| Password | Times Seen| \n| -------- | -------- | ";
                    }
                    listSummaryCommon = true;
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("| " + dr.ItemArray[0].ToString() + " | " + dr.ItemArray[1].ToString() + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + dr.ItemArray[1].ToString());
                    }
                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                if (DataTablePwnedAdmin.Rows.Count == 0)
                {
                    StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdPwned, EncodeXMLChars(StringSummaryXMLTemp));
                    TechnicalSummaryContent += StringSummaryXMLTemp;
                }

            }

            listSummaryCommon = false;
            count = 0;

            /* AdminPwned Passwords */
            foreach (DataRow dr in DataTablePwnedAdmin.Rows)
            {
                string sumString;

                if ((dr.ItemArray[1].ToString().Length) > 0 && listSummaryCommon == false)
                {
                    StringSummaryXMLTemp += ("Additionally, each of the following Administrative user accounts were identified as being compromised in previous data breaches that have been made public. The number of times that these have been seen in a data leak is listed next to the username and password: ");
                    if (TechMarkdownOutput)
                    {
                        StringSummaryXMLTemp += "\n\n|Username| Password | Times Seen| \n| -------- | -------- | --------|";
                    }
                    listSummaryCommon = true;
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("| " + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + " | " + dr.ItemArray[2].ToString() + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()) + " : " + dr.ItemArray[2].ToString());
                    }
                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(SummaryXMLVulnPwned.Text, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }

            listSummaryCommon = false;
            count = 0;

            /* Weak Passwords */
            foreach (DataRow dr in DataTableWeak.Rows)
            {
                string sumString;

                if (DataTableWeak.Rows.Count > 0 && listSummaryCommon == false)
                {
                    StringSummaryXMLTemp = ("The following user accounts and associated passwords were found to be insufficiently strong. This area of analysis is broadly generic and aims to cover the majority of the passwords that do not confirm to basic complexity rules or policy driven password requirements: ");
                    if (TechMarkdownOutput)
                    {
                        StringSummaryXMLTemp += "\n\n|Username| Password | \n| -------- | --------|";
                    }
                    listSummaryCommon = true;
                }

                if (listSummaryCommon)
                {
                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + dr.ItemArray[0].ToString() + " | " + PasswordMask(dr.ItemArray[1].ToString()) + "|");
                    }
                    else
                    {
                        sumString = ("- " + dr.ItemArray[0].ToString() + " : " + PasswordMask(dr.ItemArray[1].ToString()));
                    }
                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }
                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdWeak, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }

            listSummaryCommon = false;
            count = 0;

            /* History Repeat Password */
            foreach (DataRow dr in DataTableHistoryReuse.Rows)
            {
                string sumString;

                float historyCount = float.Parse(DataTableHistoryReuse.Rows.Count.ToString());
                float historyPercentRaw = (historyCount / totalPasswords);
                string historyPercentage = Math.Round(historyPercentRaw * 100, 0).ToString();

                if (count == 0)
                {
                    if ((int.Parse(historyCount.ToString())) > 0)
                    {
                        StringSummaryXMLTemp = ("This audit included a review of historical Active Directory passwords. These user accounts are suffixed with '_history' and are numbered incrementally from the most recent password to the oldest stored in Active Directory (for example user1, user1_history0, user1_history1, etc). The following user accounts were found to re-use a previous password as their current user account password:");
                        if (TechMarkdownOutput)
                        {
                            StringSummaryXMLTemp += "\n\n|Username| Password | \n| -------- | --------|";
                        }
                        listSummaryCommon = true;
                    }
                }

                if (listSummaryCommon)
                {


                    if (historyPercentage == "0")
                    {

                        historyPercentage = "< 1";

                    }
                    string historyPass = dr.ItemArray[1].ToString();

                    if (TechMarkdownOutput)
                    {
                        sumString = ("|" + (dr.ItemArray[0].ToString()) + "| " + PasswordMask(historyPass) + "|");
                    }
                    else
                    {
                        sumString = ("- " + (dr.ItemArray[0].ToString()) + " : " + PasswordMask(historyPass));
                    }
                    StringSummaryXMLTemp += "\n";
                    StringSummaryXMLTemp += (sumString);
                    count++;
                }

                else
                {
                    break;
                }

            }

            if (count > 0)
            {
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLTemp += "\n";
                StringSummaryXMLOutput += SummaryXMLReturnSection(Properties.Settings.Default.VulnIdHistory, EncodeXMLChars(StringSummaryXMLTemp));
                TechnicalSummaryContent += StringSummaryXMLTemp;
            }

            StringSummaryXMLOutput += "\n\t\t\t\t</vulnerabilities>\n\t\t\t</service>\n\t\t</services>\n\t</item>\n</items>";

            // Display in Technical Summary
            Dispatcher.Invoke(() => RichTextTech.Text = (TechnicalSummaryContent));

            if (CheckSummaryXML.IsChecked == true && ContinueSummaryXMLGeneration == true)
            {
                Dispatcher.Invoke(() => ButtonSummarySaveXMLFile.IsEnabled = true);
                Dispatcher.Invoke(() => ButtonSummaryRegenXMLFile.IsEnabled = true);

                Dispatcher.Invoke(() => RichTextXML.AppendText(FormatXml(StringSummaryXMLOutput)));
            }

            listSummaryCommon = false;
            count = 0;
        }

        private void SummaryOutputRecommend()
        {
            RecommendSummaryContent = "";

            int RecommendAdminCount = DataTableAdmin.Rows.Count;
            int RecommendAdminPwnedCount = DataTablePwnedAdmin.Rows.Count;
            int RecommendCommonDictCount = DataTableCommon.Rows.Count;
            int RecommendCompanyNameCount = DataTableCompany.Rows.Count;
            int RecommendReuseCount = DataTableReuse.Rows.Count;

            if (RecommendAdminCount > 0 || RecommendAdminPwnedCount > 0)
            {
                if (RecommendAdminCount > 0)
                {
                    RecommendSummaryContent += "The administrator accounts that were identified as having weak passwords " +
                        "should have their respective passwords changed and the password policy surrounding administrative or " +
                        "otherwise highly privileged user accounts should be further strengthened. As a general suggestion " +
                        "a separate password policy could be configured to require a minimum password length of 15 or more " +
                        "characters. ";
                }

                if (RecommendAdminPwnedCount > 0)
                {
                    RecommendSummaryContent += "Similarly, the administrator accounts that were found to use passwords " +
                        "that were previously seen in publicly-disclosed data-leaks should be changed. ";
                }
                RecommendSummaryContent += "\n\n";
            }

            if (RecommendCommonDictCount > 0)
            {
                RecommendSummaryContent += "Mitigating the use of commonly used passwords or those that are " +
                    "based upon plaintext dictionary words is a difficult control to implement via technical " +
                    "means. User education through security awareness training is a key recommendation " +
                    "to improve the overall security posture. However, there are third-party integrations " +
                    "that could be used to perform validation checks for keywords, dictionary lists, or industry " +
                    "recognised insecure passwords. ";
                RecommendSummaryContent += "\n\n";
            }

            if (RecommendCompanyNameCount > 0)
            {
                RecommendSummaryContent += "The use of the organisation/company name within passwords is a commonly observed practice " +
                    "with end-users. Similar to the use of dictionary-based words both user education towards " +
                    "stronger password security and technical controls can be implemented, with the latter using the " +
                    "organisation name as a basis for keyword validation. ";
                RecommendSummaryContent += "\n\n";
            }

            if (RecommendReuseCount > 0)
            {
                RecommendSummaryContent += "User accounts that utilise privilege separation are often designated via " +
                    "human-readable suffixes, for example \"user1\" and \"user1_da\". Where the identified accounts were " +
                    "found to share a password with a similarly named account it is highly recommended to enforce a " +
                    "stronger password policy for the higher-privileged accounts as a technical mitigation. ";
                RecommendSummaryContent += "\n\n";
            }

            Dispatcher.Invoke(() => RichTextRecommend.Text = (RecommendSummaryContent));
        }

        /*private void WaitForAnalysis()
        {
            MessageBox.Show(ButtonRunAnalysis.Background.ToString());
            ButtonRunAnalysis.Background = Brushes.DimGray;
            ButtonRunAnalysis.Content = "Cancel Analysis";


            ButtonRunAnalysis.Background = Brushes.DimGray;
            ButtonRunAnalysis.Content = "Run Analysis";
        }*/

        private void ClearStatusAndTime()
        {
            Dispatcher.Invoke(() => LabelStatusCharacter.Text = (""));
            Dispatcher.Invoke(() => LabelStatusHashcat.Text = (""));
            Dispatcher.Invoke(() => LabelStatusCommon.Text = (""));
            Dispatcher.Invoke(() => LabelStatusFreq.Text = (""));
            Dispatcher.Invoke(() => LabelStatusAdmin.Text = (""));
            Dispatcher.Invoke(() => LabelStatusCompany.Text = (""));
            Dispatcher.Invoke(() => LabelStatusLength.Text = (""));
            Dispatcher.Invoke(() => LabelStatusHistory.Text = (""));
            Dispatcher.Invoke(() => LabelStatusReuse.Text = (""));
            Dispatcher.Invoke(() => LabelStatusUsername.Text = (""));
            Dispatcher.Invoke(() => LabelStatusKeyboard.Text = (""));
            Dispatcher.Invoke(() => LabelStatusDate.Text = (""));

            Dispatcher.Invoke(() => LabelTimeCharacter.Text = (""));
            Dispatcher.Invoke(() => LabelTimeHashcat.Text = (""));
            Dispatcher.Invoke(() => LabelTimeCommon.Text = (""));
            Dispatcher.Invoke(() => LabelTimeFreq.Text = (""));
            Dispatcher.Invoke(() => LabelTimeAdmin.Text = (""));
            Dispatcher.Invoke(() => LabelTimeCompany.Text = (""));
            Dispatcher.Invoke(() => LabelTimeLength.Text = (""));
            Dispatcher.Invoke(() => LabelTimeHistory.Text = (""));
            Dispatcher.Invoke(() => LabelTimeReuse.Text = (""));
            Dispatcher.Invoke(() => LabelTimeUsername.Text = (""));
            Dispatcher.Invoke(() => LabelTimeKeyboard.Text = (""));
            Dispatcher.Invoke(() => LabelTimeDate.Text = (""));
        }

        public string SummaryXMLReturnSection(string vulnId, string vulnInfo)
        {
            string StringSummaryXMLTags = "\n\t\t\t\t\t<vulnerability id=\"" + vulnId + "\">\n\t\t\t\t\t\t<information>\n" + vulnInfo + "\n\t\t\t\t\t\t</information>\n\t\t\t\t\t</vulnerability>";
            return (StringSummaryXMLTags);
        }

        // Encode input XML and return encoded string
        private string EncodeXMLChars(string xmlInput)
        {
            string outPass = System.Security.SecurityElement.Escape(xmlInput);
            return (outPass);
        }

        private void ConfirmXMLOutputSettingsMissing()
        {
            if (CheckSummaryXML.IsChecked == true)
            {


                if ((CheckSummaryXML.IsChecked.Value) && ((Properties.Settings.Default.VulnIdAdmins.Length == 0) || (Properties.Settings.Default.VulnIdCharFreq.Length == 0) || (Properties.Settings.Default.VulnIdCompany.Length == 0) ||
                   (Properties.Settings.Default.VulnIdDate.Length == 0) || (Properties.Settings.Default.VulnIdDict.Length == 0) || (Properties.Settings.Default.VulnIdKeyboard.Length == 0) || (Properties.Settings.Default.VulnIdLength.Length == 0)
                   || (Properties.Settings.Default.VulnIdOverview.Length == 0) || (Properties.Settings.Default.VulnIdPwned.Length == 0) || (Properties.Settings.Default.VulnIdReuse.Length == 0) || (Properties.Settings.Default.VulnIdSource.Length == 0)
                   || (Properties.Settings.Default.VulnIdUsername.Length == 0) || (Properties.Settings.Default.VulnIdWeak.Length == 0)))
                {
                    MessageBoxResult result = MessageBox.Show("The Vulnerability IDs for the XML output have not been saved. Continue with XML generation?", "Confirmation", MessageBoxButton.YesNo);


                    if (result == MessageBoxResult.Yes)
                    {
                        ContinueSummaryXMLGeneration = true;
                    }
                    else if (result == MessageBoxResult.No)
                    {
                        ContinueSummaryXMLGeneration = false;
                    }

                }
                else if (SummaryXMLDomain.Text.Length == 0)
                {
                    var res = new PwdlyserXMLDomainPromptWindow();
                    if (res.ShowDialog() == true)
                    {
                        if (res.ResponseText.Length > 0)
                        {


                            SummaryXMLDomain.Text = res.ResponseText;
                            ContinueSummaryXMLGeneration = true;
                        }
                    }

                    //XML
                }

                else
                {
                    ContinueSummaryXMLGeneration = true;
                }
            }
        }

        // Pretty-print XML
        string FormatXml(string xml)
        {
            try
            {
                XDocument doc = XDocument.Parse(xml);
                return doc.ToString();
            }
            catch (Exception)
            {
                return xml;
            }
        }

        public IEnumerable<DataGridRow> GetDataGridRows(DataGrid grid)
        {
            var itemsSource = grid.ItemsSource as IEnumerable;
            if (null == itemsSource) yield return null;
            foreach (var item in itemsSource)
            {
                if (grid.ItemContainerGenerator.ContainerFromItem(item) is DataGridRow row) yield return row;
            }
        }

        /// <summary>
        /// Split input file to user:pass or user:hash:pass
        /// </summary>
        public bool DelimitData(string ImportFilePath)
        {
            String[] splitLine;
            string[] RawData;
            try
            {
                RawData = File.ReadAllLines(ImportFilePath, Encoding.GetEncoding("ISO-8859-1"));
            }
            catch
            {
                MessageBox.Show("Please select an input file.", "Warning: No Passwords To Anaylse");
                return false;
            }
            CreateRawDataTable();
            try
            {
                DataTableCommon.Reset();
            }
            catch
            {
                Console.WriteLine("Not Resetting Common Datatable.");
            }
            CreateDataTables();



            char[] DelimiterChars = { ':' };
            bool delimitHash;
            string username;
            string password;

            if (RadioUserPass.IsChecked.Value)
            {

                delimitHash = false;

            }
            else
            {
                delimitHash = true;
            }

            bool filterEnabled = false;

            if (RadioEnabled.IsChecked == true)
            {
                RetrieveUserDetailsForAnalysis();

                if (ADUsers.Rows.Count == 0)
                {
                    MessageBox.Show("Please use valid credentials and a target domain/IP address for Active Directory before running analysis with the 'Enabled' user account filter.");
                    RadioAll.IsChecked = true;
                }
                else
                {
                    filterEnabled = true;
                }

            }
            else
            {
                RetrieveUserDetailsForAnalysis();

            }

            if (ImportFilePath.Length > 0)
            {
                //try
                //{
                foreach (string line in RawData)
                {
                    string resultString = "";
                    if (delimitHash)
                    {
                        try
                        {


                            splitLine = line.Split(DelimiterChars, 3);
                            username = splitLine[0];
                            password = splitLine[2];
                        }
                        catch
                        {
                            MessageBox.Show("Unable to split usernames and passwords. Please manually check the input file.\n\nThefollowing formats are accepted:\n- username:password\n- username:LM/NTLM hash:password", "Error: Input Format");
                            return false;
                        }
                    }
                    else
                    {
                        try
                        {
                            splitLine = line.Split(DelimiterChars, 2);
                            username = splitLine[0];
                            password = splitLine[1];
                        }

                        catch
                        {
                            MessageBox.Show("Unable to split usernames and passwords. Please manually check the input file.\n\nThefollowing formats are accepted:\n- username:password\n- username:LM/NTLM hash:password", "Error: Input Format");
                            return false;
                        }
                    }

                    if (username != string.Empty)
                    {
                        if (password.Length > 4)
                        {

                            if (password.Substring(0, 5) == "$HEX[")
                            {
                                resultString = password.Split('[', ']')[1];
                                password = (HextoString(resultString).ToString());
                            }

                        }

                        if (filterEnabled)
                        {
                            if (ADUsers.Rows.Count > 0)
                            {
                                string tempusername = username;

                                if (username.Contains('\\'))
                                {
                                    tempusername = (username.Split('\\')[1]);
                                }
                                var results = from DataRow myRow in ADUsers.Rows
                                              where (string)myRow["Username"].ToString().ToLower() == tempusername.ToLower()
                                              select myRow;


                                if (results.Count() > 0)
                                {
                                    if (results.ToArray()[0].ItemArray[3].ToString() == "Enabled")
                                    {
                                        RawDataTable.Rows.Add(username, password);
                                    }
                                }

                            }
                            else
                            {
                                RawDataTable.Rows.Add(username, password);
                            }

                        }
                        else
                        {
                            RawDataTable.Rows.Add(username, password);
                        }

                    }
                }
            }

            RawDataTableBackup = RawDataTable;

            return true;
        }

        // Decode Hashcat HEX[] strings to plaintext passwords
        public static string HextoString(string InputText)
        {

            byte[] bb = Enumerable.Range(0, InputText.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(InputText.Substring(x, 2), 16))
                             .ToArray();
            return Encoding.GetEncoding("ISO-8859-1").GetString(bb);
        }

        private string DateTimePass(string password)
        {

            var dateTimeList = new List<string>
            {
                    {"january"},
                    {"february"},
                    {"march"},
                    {"april"},
                    {"june"},
                    {"july"},
                    {"august"},
                    {"september"},
                    {"october"},
                    {"november"},
                    {"december"},
                    {"janvier"},
                    {"fevrier"},
                    //{"mars"},
                    //{"avril"},
                    //{"mai"},
                    {"juin"},
                    {"juillet"},
                    {"aout"},
                    {"septembre"},
                    {"octobre"},
                    {"novembre"},
                    {"decembre"},
                    {"enero"},
                    {"febrero"},
                    {"marzo"},
                    {"abril"},
                    //{"mayo"},
                    {"junio"},
                    {"julio"},
                    {"agosto"},
                    {"septiembre"},
                    {"octubre"},
                    {"noviembre"},
                    {"diciembre"},
                    {"monday"},
                    {"tuesday"},
                    {"wednesday"},
                    {"thursday"},
                    {"friday"},
                    {"saturday"},
                    {"sunday"},
                    {"spring"},
                    {"summer"},
                    {"autumn"},
                    {"winter"},
            };

            string b = dateTimeList.FirstOrDefault(dateTimePass => Deleet_password(password).ToLower().Contains(dateTimePass.ToLower().Replace("\r", "")));
            if (b != null)
            {
                return b.Replace("\r", "");
            }
            else
            {
                return ("NOTVALID");
            }


        }

        public string FirstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToUpper(str[0]) + str.Substring(1);

            return str.ToUpper();
        }

        // Mask password input characters
        static string Mask(int n)
        {
            return new String('*', n);
        }

        // Update ADUsers table to align with known cracked-passwords
        private void ADUsersUpdatedCracked()
        {
            if (ADUsers.Rows.Count > 0)
            {

                foreach (DataRow rawRow in RawDataTable.Rows)
                {
                    string username = rawRow[0].ToString();

                    if (username.Contains('\\'))
                    {
                        username = username.Split(new char[] { '\\' })[1];
                    }
                    foreach (DataRow adRow in ADUsers.Rows)
                    {
                        // Set from raw data
                        if (adRow["Username"].ToString().ToLower() == username.ToLower())
                        {
                            adRow["Cracked"] = true;
                        }

                    }
                }
                Dispatcher.Invoke(() => DataGridUserProperties.ItemsSource = ADUsers.AsDataView());
                Dispatcher.Invoke(() => DataGridUserProperties.Items.Refresh());

            }
        }

        // Mask passwords varying length - return masked password
        private string PasswordMask(string pwd)
        {
            string mask_pwd = pwd;

            if (CheckSummaryMask.IsChecked.Value)
            {
                mask_pwd = pwd;
                if (pwd.Length == 0)
                {
                    mask_pwd = "*****BLANK PASSWORD*****";
                }
                else
                {
                    if (pwd.Length > 9)
                    {
                        mask_pwd = pwd.Substring(0, 3) + Mask((pwd.Length - 6)) + pwd.Substring((pwd.Length - 3)); // -3
                    }
                    else if (pwd.Length >= 5 && pwd.Length <= 9)
                    {
                        mask_pwd = pwd.Substring(0, 2) + Mask((pwd.Length - 4)) + pwd.Substring(pwd.Length - 2); //-2
                    }
                    else if (pwd.Length == 4)
                    {
                        mask_pwd = pwd.Substring(0, 1) + Mask((pwd.Length - 2)) + pwd.Substring(pwd.Length - 1); //-1
                    }
                    else if (pwd.Length == 3)
                    {
                        mask_pwd = pwd.Substring(0, 1) + Mask((pwd.Length - 2)) + pwd.Substring(pwd.Length - 1); //-1
                    }
                    else if (pwd.Length == 2)
                    {
                        mask_pwd = pwd.Substring(0, 1) + Mask((pwd.Length - 1));
                    }
                }
            }

            return (mask_pwd);
        }

        // Check password against known keyboard patterns
        private string KeyboardPass(string password)
        {
            var keyboardList = new List<string>
            {
                {"1qaz2wsx"},
                {"!qaz2wsx"},
                {"!qaz\"wsx"},
                {"qwer1234"},
                {"a1z2e3r4t5y6"},
                {"1a2z3e4r5t6y"},
                {"lkjhgfdsa"},
                {"asdfghjkl"},
                {"asdfgh"},
                {"poiuytrewq"},
                {"qwertyuiop"},
                {"qwerty"},
                {"q1w2e3"},
                {"z1x2c3"},
                {"1q2w3e"},
                {"nkoplm"},
                {"zaqwsx"},
                {"qazwsx"},
                {"qazxsw"},
                {"zxcvbn"},
                {"zxcdsa"},
                {"zaqxsw"},
                {"azerty"},
                {"dvorak"},
                {"plmnko"},
                {"mnbhjk"},
                {"mnbvc"},
                {"hjkl"},
                {"azert"},
                {"asdf"},
                {"lkjh"},
                {"qwer"},
                {"zxc"},
                {"1qaz"},
                {"2wsx"},
                {"poiu"},
                {"plm"},
                {"2468"},
                {"1357"},
                {"3579"},
                {"0864"},
            };

            string b = keyboardList.FirstOrDefault(keyboardPass => Deleet_password(password).ToLower().Contains(keyboardPass.ToLower().Replace("\r", "")));
            if (b != null)
            {
                return b.Replace("\r", "");
            }
            else
            {
                return ("NOTVALID");
            }

            //foreach (string keyboardPass in keyboardList)
            //{
            //    string deleeted = Deleet_password(password);
            //    if (deleeted.ToString().ToLower().Contains(keyboardPass.ToLower()))
            //    {
            //        return (keyboardPass.ToString());
            //    }
            //}
            //return ("NOTVALID");
        }

        public class WebClientWithTimeout : WebClient
        {
            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest wr = base.GetWebRequest(address);
                wr.Timeout = 4000; // timeout in milliseconds (ms)
                return wr;
            }
        }

        // HaveIBeenPwned API request constructor
        private string PwnedAPIRequest(string hashRange)
        {
            WebClient client = new WebClientWithTimeout();
            client.Headers.Add("user-agent", "Pwnage-Checker-Windows-Pwdlyser");

            try
            {
                Stream stream = client.OpenRead("https://api.pwnedpasswords.com/range/" + hashRange);
                client.Dispose();

                StreamReader reader = new StreamReader(stream);
                String content = reader.ReadToEnd();
                return content;
            }
            catch (WebException e)
            {
                client.Dispose();

                if (PwnedAPIErrorResponse == false)
                {
                    MessageBox.Show(e.Message, "Error: Passwords in Data Leaks");
                    PwnedAPIErrorResponse = true;
                }


                return "ERROR";
            }


        }

        // Hash SHA1 password for HaveIBeenPwned check
        static string Sha1Hash(string input)
        {
            var hash = (new System.Security.Cryptography.SHA1Managed()).ComputeHash(Encoding.UTF8.GetBytes(input));
            return string.Join("", hash.Select(b => b.ToString("x2")).ToArray());
        }

        // User specified max HaveIBeenPwned checks
        public void SetMaxPwnedCheck()
        {
            maxpwncheck = int.Parse(TextBoxPwnedItems.Text);
            if (maxpwncheck > 1000)
            {
                MessageBox.Show("Please note that checking " + maxpwncheck.ToString() + " passwords may take some time.");
            }

        }

        public void CheckAdminPwnedBoolSet()
        {
            if (ListBoxAdmin.Items.Count > 0)
            {
                PwnedAdminCheckBool = true;
            }
            else
            {
                PwnedAdminCheckBool = false;
            }
        }

        // Convert Extracted Suffix to Hashcat rule format
        private string HashcatSuffixRule(string password)
        {
            string rule = "";

            foreach (char c in password)
            {
                rule += "$" + c.ToString();
            }
            return (rule);
        }

        public static string ReverseString(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }



        // Change back to base words
        private string Deleet_password(string password)
        {
            string[,] deleetList = new string[,]
{
                {"4,a"},
                {"1,i"},
                {"3,e"},
                {"!,i"},
                {"5,s"},
                {"6,g"},
                {"0,o"},
                {"$,s"},
                {"7,t"},
                {"9,g"},
                {"@,a"},
                {"<,c"},
                {"£,e"},
                {"|,l"},
                {"8,b"},
                {"(,c"},
                {"+,t"},
                {";,l"},
                {":,l"}

            };

            string deleetedPass = password;
            if (deleetedPass.Length == 0)
            {
                return (deleetedPass.ToString());
            }

            foreach (string charStart in deleetList)
            {
                string[] changeChar = charStart.Split(',');
                deleetedPass = deleetedPass.Replace(changeChar[0].ToString(), changeChar[1].ToString());
            }

            return (deleetedPass.ToString());
        }

        // Analysis - Import password list - return file path
        public string ImportUserPassList()
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                DefaultExt = ".txt",
                Filter = "Documents|*.txt; *.lc; *.hash|All|*.*"
            };

            bool? result = dlg.ShowDialog();

            // Set text box for file path
            ImportFilePath = dlg.FileName;
            if (ImportFilePath.Length > 0)
            {
                TabSummary.Visibility = Visibility.Collapsed;
            }

            return ImportFilePath;
        }

        // Common/Dict - Set base dictionary list
        public void SetLangDictResource()
        {
            langdict = Properties.Resources.langdict.ToString();
            langdictarray = langdict.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

            langdictlarge = Properties.Resources.langdictlarge.ToString();
            langdictlargearray = langdictlarge.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
        }

        private string CommonPass(string password)
        {

                if (commondictcheck == true)
                {
                    if (commondictchecklarge)
                    {
                    // loops through large lang dictionary
                    //foreach (string commonPass in langdictlargearray)
                    string deleeted = Deleet_password(password).ToLower();
                    var b = langdictlargearray.AsParallel().FirstOrDefault(s => deleeted.Contains(s.ToLower().Replace("\r", "")));

                    if (b != null)
                    {
                        return b.Replace("\r","");
                    }
                    else
                    {
                        return ("NOTVALID");
                    }

                    }

                else
                {
                        // loops through large lang dictionary
                        //foreach (string commonPass in langdictlargearray)
                        string deleeted = Deleet_password(password).ToLower();
                        var b = langdictarray.AsParallel().FirstOrDefault(s => deleeted.Contains(s.ToLower().Replace("\r", "")));

                        if (b != null)
                        {
                            return b.Replace("\r", "");
                        }
                        else
                        {
                            return ("NOTVALID");
                        }

                }
            }

            else
            {
                if (CommonPassListArray.Count() > 0)
                {

                
                    string deleeted = Deleet_password(password).ToLower();
                    var b = CommonPassListArray.AsParallel().FirstOrDefault(s => deleeted.Contains(s.ToLower().Replace("\r", "")));

                    if (b != null)
                    {
                        return b.Replace("\r", "");
                    }
                    else
                    {
                        return ("NOTVALID");
                    }
                }
                else
                {
                    return ("NOTVALID");
                }

            }


        }

        public DataView HistorySelect(string user)
        {
            var historyRegex = new Regex(@"_history[0-9]");

            DataTable DataTableHistorySelect = new DataTable();
            DataTableHistorySelect.Columns.AddRange(new[]
                                       {
                               new DataColumn("Usernames", typeof(string)),
                               new DataColumn("Passwords", typeof(string)),
                });

            foreach (DataRow row in RawDataTable.Rows)
            {

                string user1 = row[0].ToString();
                string pass1 = row[1].ToString();


                if (user1.Split('_')[0] == (user))
                {
                    if (historyRegex.Match(user1).Success)
                    {
                        DataTableHistorySelect.Rows.Add(user1, pass1);
                    }
                }
            }
            DataView dvChar = DataTableHistorySelect.DefaultView;
            dvChar.Sort = "Usernames asc";
            DataTable sortedDTChar = dvChar.ToTable();

            return sortedDTChar.AsDataView();
        }

        private string CreateWordList()
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                DefaultExt = ".txt",
                Filter = "Documents|*.txt;|All|*.*"
            };

            bool? result = dlg.ShowDialog();

            // Set text box for file path
            ImportFilePath = dlg.FileName;
            return ImportFilePath;
        }

        private string RemoveTrailingNumbers(string password)
        {

            string pattern = @"\d+$";
            string replacement = "";
            Regex rgx = new Regex(pattern);
            Regex rgxAllNumbers = new Regex(@"^[0-9]+$");

            string result = rgx.Replace(password, replacement);

            return (result.ToString());
        }

        // Wordlist generator - remove special characters
        private string RemoveSpecial(string password)
        {
            string[,] specialList = new string[,]
            {
                {" "},
                {"~"},
                {"}"},
                {"|"},
                {"{"},
                {"`"},
                {"_"},
                {"^"},
                {"]"},
                {"\\"},
                {"["},
                {"@"},
                {"?"},
                {">"},
                {"="},
                {"<"},
                {";"},
                {":"},
                {"/"},
                {"."},
                {"-"},
                {","},
                {"+"},
                {"*"},
                {")"},
                {"("},
                {"'"},
                {"&"},
                {"%"},
                {"$"},
                {"#"},
                {"\""},
                {"!"},
                {"_"},
            };

            string result = "";

            foreach (string specChar in specialList)
            {

                if (password.Contains(specChar))
                {
                    result = password.Replace(specChar, "");
                }
            }

            return (result.ToString());
        }

        // Wordlist generator - remove trailing special characters
        private string RemoveTrailingSpecial(string password)
        {
            string filtered = password;
            filtered = Regex.Replace(filtered, @"\W+", "");   // trim non-word characters
            filtered = Regex.Replace(filtered, @"^\d+", "");  // remove number(s) at start

            return (filtered.ToString());
        }

        // Wordlist generator - convert input list to 'cleaned' wordlist
        private void ProcessWordList(string[] wordlist)
        {
            CreateWordlistDataTable();
            Dispatcher.Invoke(() => CreateWordlist.IsEnabled = false);
            Dispatcher.Invoke(() => CreateWordlistAnalysis.IsEnabled = false);

            bool original = Dispatcher.Invoke(() => CheckBoxWordlistOriginal.IsChecked.Value);
            bool deleet = Dispatcher.Invoke(() => CheckBoxWordlistDeleet.IsChecked.Value);
            bool numbers = Dispatcher.Invoke(() => CheckBoxWordlistNumbers.IsChecked.Value);
            bool special = Dispatcher.Invoke(() => CheckBoxWordlistSpecial.IsChecked.Value);
            bool specialAll = Dispatcher.Invoke(() => CheckBoxWordlistSpecialAll.IsChecked.Value);
            bool freqSuffix = Dispatcher.Invoke(() => CheckBoxWordlistFreqSuffix.IsChecked.Value);


            bool lowercase = Dispatcher.Invoke(() => CheckBoxWordlistLowCase.IsChecked.Value);
            bool lowercaseCamel = Dispatcher.Invoke(() => CheckBoxWordlistCamel.IsChecked.Value);
            bool suffixString = Dispatcher.Invoke(() => CheckBoxWordlistSuffix.IsChecked.Value);
            bool prefixString = Dispatcher.Invoke(() => CheckBoxWordlistPrefix.IsChecked.Value);
            string prefix = Dispatcher.Invoke(() => TextBoxWordlistPrefix.Text.ToString());
            string suffix = Dispatcher.Invoke(() => TextBoxWordlistSuffix.Text.ToString());
            bool recursive = Dispatcher.Invoke(() => CheckBoxWordlistRecursive.IsChecked.Value);

            List<string> list_lines = new List<string>(wordlist);

            foreach (string line in wordlist)
            {
                // Keep original value  
                if (original)
                {
                    DataTableWordList.Rows.Add(line);
                }

                // De-leet password

                if (deleet)
                {
                    DataTableWordList.Rows.Add(Deleet_password(line));
                }

                // Remove trailing number
                if (numbers)
                {
                    DataTableWordList.Rows.Add(RemoveTrailingNumbers(line));
                }

                // Remove trailing special and numbers
                if (special)
                {
                    DataTableWordList.Rows.Add(RemoveTrailingSpecial(line));
                }

                // Remove all special
                if (specialAll)
                {
                    DataTableWordList.Rows.Add(RemoveSpecial(line));
                }

                // Lowercase all
                if (lowercase)
                {
                    DataTableWordList.Rows.Add(Deleet_password(RemoveTrailingNumbers(line)).ToLower());
                }

                // First letter to upper
                if (lowercaseCamel)
                {
                    DataTableWordList.Rows.Add(FirstLetterToUpper(Deleet_password(RemoveTrailingNumbers(line)).ToLower()));
                }

                // Prefix String
                if (prefixString)
                {
                    DataTableWordList.Rows.Add(prefix + line);

                }

                // Prefix String
                if (suffixString)
                {
                    DataTableWordList.Rows.Add(line + suffix);

                }

                // Prefix String
                if (freqSuffix)
                {
                    int SuffixCount = 0;
                    foreach (DataRow row in DataTableTrailingSuffix.AsEnumerable())
                    {
                        if (SuffixCount < 20)
                        {
                            string trailingSuffix = row[0].ToString();
                            DataTableWordList.Rows.Add(line + trailingSuffix);
                            SuffixCount++;
                        }
                        else
                        {
                            break;
                        }
                    }
                }

                if (recursive)
                {

                    // Remove all special
                    if (specialAll && deleet)
                    {
                        DataTableWordList.Rows.Add(Deleet_password(RemoveSpecial(line)));
                    }

                    // Remove special and deleet
                    if (special && deleet)
                    {
                        DataTableWordList.Rows.Add(Deleet_password(RemoveTrailingSpecial(line)));
                    }

                    // Remove trailing number and deleet
                    if (numbers && deleet)
                    {
                        DataTableWordList.Rows.Add(Deleet_password(RemoveTrailingNumbers(line)));
                    }

                    // Prefix String && Deleet
                    if (prefixString && deleet)
                    {
                        DataTableWordList.Rows.Add(prefix + Deleet_password(line));
                    }
                    // Prefix String && Lower
                    if (prefixString && lowercase)
                    {
                        DataTableWordList.Rows.Add(prefix + (line.ToLower()));
                    }
                    // Prefix String && Lower Camel
                    if (prefixString && lowercaseCamel)
                    {
                        DataTableWordList.Rows.Add(prefix + FirstLetterToUpper(line.ToLower()));
                    }
                    // Prefix String && Lower && deleet
                    if (prefixString && lowercase && deleet)
                    {
                        DataTableWordList.Rows.Add(prefix + Deleet_password(line.ToLower()));
                    }
                    // Prefix String && Lower Camel && deleet
                    if (prefixString && lowercaseCamel && deleet)
                    {
                        DataTableWordList.Rows.Add(prefix + FirstLetterToUpper(Deleet_password(line)));
                    }
                    // Prefix String && remove trailing number && deleet
                    if (prefixString && lowercase && deleet)
                    {
                        DataTableWordList.Rows.Add(prefix + Deleet_password(RemoveTrailingNumbers(line.ToLower())));
                    }

                    // suffix String && Deleet
                    if (suffixString && deleet)
                    {
                        DataTableWordList.Rows.Add(Deleet_password(line) + suffix);
                    }
                    // suffix String && Lower
                    if (suffixString && lowercase)
                    {
                        DataTableWordList.Rows.Add((line.ToLower()) + suffix);
                    }
                    // suffix String && Lower Camel
                    if (suffixString && lowercaseCamel)
                    {
                        DataTableWordList.Rows.Add(FirstLetterToUpper(line.ToLower()) + suffix);
                    }
                    // suffix String && Lower && deleet
                    if (suffixString && lowercase && deleet)
                    {
                        DataTableWordList.Rows.Add(Deleet_password(line.ToLower()) + suffix);
                    }
                    // suffix String && Lower Camel && deleet
                    if (suffixString && lowercaseCamel && deleet)
                    {
                        DataTableWordList.Rows.Add(FirstLetterToUpper(Deleet_password(line)) + suffix);
                    }
                    // suffix String && remove trailing number && deleet
                    if (suffixString && lowercase && deleet)
                    {
                        DataTableWordList.Rows.Add(Deleet_password(RemoveTrailingNumbers(line.ToLower())) + suffix);
                    }

                }
            }

            DataTableWordListDistinct = DataTableWordList.DefaultView.ToTable(true, "Passwords");
            Dispatcher.Invoke(() => DataGridWordlist.ItemsSource = DataTableWordListDistinct.AsDataView());
            Dispatcher.Invoke(() => CreateWordlist.IsEnabled = true);
            try
            {
                if (RawDataTable.Rows.Count > 0)
                {
                    Dispatcher.Invoke(() => CreateWordlistAnalysis.IsEnabled = true);
                }
            }
            catch
            {
                Dispatcher.Invoke(() => CreateWordlistAnalysis.IsEnabled = false);
            }

            Dispatcher.Invoke(() => SaveWordlist.IsEnabled = true);
        }

        // Wordlist generator - Save wordlist to file
        private void WriteWordListToFile(string filepath)
        {
            StreamWriter objWriter = new StreamWriter(filepath, false);
            foreach (DataRow row in DataTableWordListDistinct.Rows)
            {
                objWriter.WriteLine(row.ItemArray[0].ToString());
            }
            objWriter.Close();
        }

        // Wordlist generator - Load wordlist
        private void LoadBaseWordList(object sender, RoutedEventArgs e)
        {
            if (ListBoxBaseword.Items.Count == 0)
            {

                string[,] commonList = new string[,]
                {
                {"password"},
                {"123456"},
                {"letmein"},
                {"system"},
                {"admin"},
                {"starwars"},
                {"football"},
                {"welcome"},
                {"abc123"},
                {"dragon"},
                {"iloveyou"},
                {"monkey"},
                {"login"},
                {"soccer"},
                {"freedom"},
                {"love"},
                {"contraseña"}, //Spanish - Password
                {"contrasena"}, //Spanish - Password
                {"bienvenido"}, // Spanish - Welcome
                {"déjameentrar"}, // Spanish - Let me in
                {"dejameentrar"}, // Spanish - Let me in
                {"sistema"}, // Spanish - System
                {"bienvenue"}, // French - Welcome
                {"système"}, // French - system
                {"systeme"}, // French - system
                {"laissemoientrer"}, // French -Let Me in
                {"welkom"}, // german - welcome
                {"motdepasse"}, // french - password
                {"passwort"}, // Ditch - password
                {"wachtwoord"}, // dutch - password
                {"willkommen" }, // german - welcome
                {"systeem"}, // dutch - system
                {"laatmebinnen" }, // dutch - letmein
                {"lassmichrein" }, // german - letmein

                };

                foreach (string commonPass in commonList)
                {
                    ListBoxBaseword.Items.Add(commonPass);
                }
            }
        }

        // Textbox control - only permit INTs
        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        public void CreateTotalXML()
        {
            var filePath = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.TotalOverTimeSettingsPath);

            // Hacky job to create initial Pwdlyser directory in app-data. Not being created by default for some reason?
            try
            {
                Directory.CreateDirectory(filePath.Replace("pwdlyser-total.xml",""));
            }
            catch
            {

            }

            using (StreamWriter sw = File.CreateText(filePath))
            {
                sw.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sw.WriteLine("<DocumentElement>");
                sw.WriteLine("<Item date=\"1970-01-01\" total =\"0\"/>");
                sw.WriteLine("</DocumentElement>");

            }
        }

        public void LoadTotalXMLFile()
        {
            DataSet ds = new DataSet();
            var filePath = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.TotalOverTimeSettingsPath);


            ds.ReadXml(filePath, XmlReadMode.InferSchema);
            DataGridTotalCracked.ItemsSource = ds.Tables[0].DefaultView;
            Dispatcher.Invoke(() => GenerateTotalOverTime());

        }

        public void AddToProjectXML()
        {
            DataTable dt = new DataTable();

            dt = ((DataView)DataGridTotalCracked.ItemsSource).ToTable();

            if (dt.Rows.Count == 1)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string tableDate = dr.ItemArray[0].ToString();
                    if (tableDate == "1970-01-01")
                    {
                        dt.Rows.Remove(dt.Rows[0]);
                        break;
                    }
                }
            }

            string a = TextBoxTotalManualAdd.Text;
            string b = TextBoxTotalManualAddDate.Text;

            if (DataGridUserProperties.Items.Count >= RawDataTable.Rows.Count)
            {
                float tempDomainTotalUsers = ADUsers.Rows.Count;

                if (CheckADEnabled.IsChecked == true)
                {
                    DataTable tblFiltered = ADUsers.AsEnumerable()
                     .Where(r => r.Field<string>("Account Status") == "Enabled")
                     .CopyToDataTable();
                    tempDomainTotalUsers = tblFiltered.Rows.Count;
                    tblFiltered.Dispose();
                    MessageBox.Show("Please note that Analysis Over Time has been updated to only include 'Enabled' user accounts.", "Warning: Analysis Over Time");

                }
                else if (CheckADDisabled.IsChecked == true)
                {
                    DataTable tblFiltered = ADUsers.AsEnumerable()
                     .Where(r => r.Field<string>("Account Status") == "Disabled")
                     .CopyToDataTable();
                    tempDomainTotalUsers = tblFiltered.Rows.Count;
                    tblFiltered.Dispose();
                    MessageBox.Show("Please note that Analysis Over Time has been updated to only include 'Disabled' user accounts.", "Warning: Analysis Over Time");
                }
                float tempDomainCrackedUsers = RawDataTable.Rows.Count;
                float tempDomainPercentageCracked = (tempDomainCrackedUsers / tempDomainTotalUsers) * 100;
                a = (Math.Round(tempDomainPercentageCracked, 0).ToString());
            }

            string[] total = { b, a };

            dt.Rows.Add(total);

            DataGridTotalCracked.ItemsSource = dt.AsDataView();
            ChartTotalYAxis.MaxValue = 101;
            Dispatcher.Invoke(() => GenerateTotalOverTime());


            ExportDgvToXML();
        }

        private void ExportDgvToXML()
        {
            var filePath = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.TotalOverTimeSettingsPath);

            DataTable dt = new DataTable();
            dt = ((DataView)DataGridTotalCracked.ItemsSource).ToTable();

            dt.WriteXml(filePath);
        }

        internal void CreateAdminsTable()
        {


            ADAdmins = new DataTable();
            ADAdmins.Columns.AddRange(new[]
                              {
                            new DataColumn("SamAccountName", typeof(string)),
                            new DataColumn("Name", typeof(string)),

                });
        }

        private void addAdmins()
        {
            CreateAdminsTable();

            DirectoryEntry myLdapConnection = createDirectoryEntry();
            DirectorySearcher search = new DirectorySearcher(myLdapConnection);
            search.Filter = "(&(ObjectClass=Group)(cn=Administrators))";
            SearchResult builtIn = search.FindOne();
            search.PropertiesToLoad.Add("distinguishedname");
            string builtInAdminDN = builtIn.Properties["distinguishedname"][0].ToString();

            search.Filter = "(&(objectClass=user)(memberof:1.2.840.113556.1.4.1941:=" + builtInAdminDN + "))";

            search.PropertiesToLoad.Add("CN");
            search.PropertiesToLoad.Add("objectcategory");
            search.PropertiesToLoad.Add("samaccountname");

            SearchResultCollection results = search.FindAll();
            if (results != null) // change to result
            {
                foreach (SearchResult result in results)
                {

                    DataRow row = ADAdmins.NewRow();

                    string ADgroupname = result.Properties["CN"][0].ToString();


                    try
                    {
                        string samaccountname = result.Properties["samaccountname"][0].ToString();

                        row[0] = samaccountname;
                        row[1] = ADgroupname;
                        //row[2] = ADDescription;
                        //row[3] = distinguishedname;
                        ADAdmins.Rows.Add(row);
                    }
                    catch
                    {
                        continue;
                    }
                }
                DataGridAdminGroups.ItemsSource = ADAdmins.AsDataView();
            }
        }

        private DirectoryEntry createDirectoryEntry()

        {
            string domain = TextBoxDomainName.Text;
            string username = TextBoxADUsername.Text;
            string password = TextBoxADPassword.Password;
            if (username.Length == 0)
            {
                DirectoryEntry ldapConnection = new DirectoryEntry(domain);
                ldapConnection.Path = "LDAP://" + domain.ToString();
                return ldapConnection;
            }
            else
            {
                DirectoryEntry ldapConnection = new DirectoryEntry(domain, username, password);
                ldapConnection.Path = "LDAP://" + domain.ToString();
                ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
                return ldapConnection;
            }
        }

        public void GetADUsers()
        {
            string domain = TextBoxDomainName.Text;
            string username = TextBoxADUsername.Text;
            string password = TextBoxADPassword.Password;
            var users = ADUser.GetUsers("LDAP://" + domain, username, password);

            DataTable ADUsersTemp = new DataTable();
            ADUsersTemp.Columns.AddRange(new[]
                                       {
                               new DataColumn("Username", typeof(string)),
                               new DataColumn("Last Logon", typeof(int)),
                               new DataColumn("Days Since Last Set", typeof(int)),
                               new DataColumn("Account Status", typeof(string)),
                               new DataColumn("Description", typeof(string)),
                               new DataColumn("Cracked", typeof(bool))

                });

            ADUsersTemp.Columns["Cracked"].DefaultValue = false;

            DateTime now = DateTime.Now;

            bool adUserResults = false;

            foreach (var user in users)
            {

                long value = (long)user.PwdLastSet;
                int pwdLastSet = (int)(now - DateTime.FromFileTimeUtc(value)).TotalDays;

                long lastLogonValue = (long)user.LastLogin;
                int lastLogon = (int)(now - DateTime.FromFileTimeUtc(lastLogonValue)).TotalDays;

                if (DateTime.FromFileTimeUtc(user.LastLogin).ToString().Contains("01/01/1601"))
                {
                    lastLogon = -1;
                }
                if (DateTime.FromFileTimeUtc(user.PwdLastSet).ToString().Contains("01/01/1601"))
                {
                    pwdLastSet = -1;
                }

                if (RawDataTable.Rows.Count > 0)
                {
                    bool contains = RawDataTable.AsEnumerable().Any(row => user.SamAcountName.ToLower() == row.Field<String>("Usernames").ToLower());
                    ADUsersTemp.Rows.Add(user.SamAcountName, lastLogon, pwdLastSet, user.UserAccountControl, user.Description, contains);

                    adUserResults = true;
                }
                else
                {
                    ADUsersTemp.Rows.Add(user.SamAcountName, lastLogon, pwdLastSet, user.UserAccountControl, user.Description);
                    adUserResults = true;
                }

            }

            DataView dv = ADUsersTemp.DefaultView;
            dv.Sort = "Last Logon DESC";

            ADUsers = dv.ToTable();

            DataGridUserProperties.ItemsSource = dv;
            ADFilterGrid.Visibility = Visibility.Visible;

            if (adUserResults)
            {
                RadioAnalyseSelection.Visibility = Visibility.Visible;
            }
        }

        private void RetrieveUserDetailsForAnalysis()
        {
            if (TextBoxDomainName.Text.Length > 0)
            {

                try
                {
                    GetADUsers();
                }
                catch
                {

                }
                try
                {
                    addAdmins();
                    if (ADAdmins.Rows.Count > 0)
                    {
                        Dispatcher.Invoke(() => ListBoxAdmin.Items.Clear());
                        Dispatcher.Invoke(() => AddAdminsFromADList());
                    }
                    else
                    {
                        MessageBox.Show("No Administrator or nested group user accounts detected. Please enter manually within the Administrator analysis tab.");
                    }
                }
                catch
                {

                }
            }
        }

        private void AddAdminsFromADList()
        {
            Properties.Settings.Default.AdminAccountsList = "";
            foreach (DataRow admin in ADAdmins.AsEnumerable())
            {
                ListBoxAdmin.Items.Add(admin[0].ToString());
                Properties.Settings.Default.AdminAccountsList = admin[0].ToString() + ";" + Properties.Settings.Default.AdminAccountsList;
            }

            Properties.Settings.Default.Save();
        }

        private void ADEnableFilter()
        {
            try
            {
                DataTable tblFiltered = ADUsers.AsEnumerable()
                                 .Where(r => r.Field<string>("Account Status") == "Enabled")
                                 .CopyToDataTable();
                DataGridUserProperties.ItemsSource = tblFiltered.AsDataView();
            }
            catch
            {

            }
        }

        private void ADDisabledFilter()
        {
            try
            {


                DataTable tblFiltered = ADUsers.AsEnumerable()
                     .Where(r => r.Field<string>("Account Status") == "Disabled")
                     .CopyToDataTable();
                DataGridUserProperties.ItemsSource = tblFiltered.AsDataView();
            }
            catch { }
        }

        private void ADCrackedFilter()
        {
            try
            {
                DataTable tblFiltered = ADUsers.AsEnumerable()
                     .Where(r => r.Field<bool>("Cracked") == true)
                     .CopyToDataTable();
                DataGridUserProperties.ItemsSource = tblFiltered.AsDataView();
            }
            catch { }
        }

        private void SaveDomainData()
        {
            try
            {
                SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
                dlg.DefaultExt = ".csv";
                dlg.Filter = "Comma-Separated File (CSV)|*.csv";

                if (dlg.ShowDialog() == true)
                {
                    string filename = dlg.FileName;
                    DataGridUserProperties.SelectAllCells();

                    DataGridUserProperties.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;
                    ApplicationCommands.Copy.Execute(null, DataGridUserProperties);

                    DataGridUserProperties.UnselectAllCells();

                    string result = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);

                    File.AppendAllText(filename, result, UnicodeEncoding.UTF8);
                }
            }
            catch
            {
                MessageBox.Show("Please retrieve user details from Active Directory before saving to CSV.");
            }
        }

        private void LoadDomainData()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".csv";
            dlg.Filter = "Documents|*.csv; *.txt|All|*.*";

            if (dlg.ShowDialog() == true)
            {

                String textLine = string.Empty;
                String[] splitLine;

                // clear the grid view

                var data = new DataTable();
                data.Columns.AddRange(new[]
                                       {
                               new DataColumn("Username", typeof(string)),
                               new DataColumn("Last Logon", typeof(int)),
                               new DataColumn("Days Since Last Set", typeof(int)),
                               new DataColumn("Account Status", typeof(string)),
                               new DataColumn("Description", typeof(string)),
                               new DataColumn("Cracked", typeof(bool)),
                            });

                data.Rows.Clear();

                if (File.Exists(dlg.FileName))
                {
                    try
                    {
                        StreamReader objReader = new StreamReader(dlg.FileName);

                        var contents = objReader.ReadToEnd();

                        var strReader = new StringReader(contents);

                        do
                        {
                            textLine = strReader.ReadLine();
                            if ((textLine != string.Empty) && textLine != "Username,Last Logon,Days Since Last Set,Account Status,Description,Cracked")
                            {
                                try
                                {
                                    splitLine = textLine.Split(',');
                                    if (splitLine[0] != string.Empty)
                                    {
                                        splitLine[5] = "false";
                                        data.Rows.Add(splitLine);
                                        ADUsers.Rows.Add(splitLine);
                                    }
                                }
                                catch
                                {

                                    if (textLine != string.Empty)
                                    {

                                        data.Rows.Add(textLine, 0, 0, "", "", false);
                                        ADUsers.Rows.Add(textLine, 0, 0, "", "", false);
                                    }
                                }
                            }
                        } while (strReader.Peek() != -1);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }

                DataGridUserProperties.ItemsSource = data.AsDataView();
            }
        }

        private void DynamicSearch()
        {

            if (TextBoxSearchInput.Text.Length > 0)
            {
                if (RawDataTable.Rows.Count > 0)
                {

                    var lower_regex = new Regex(@"[a-z]");
                    var higher_regex = new Regex(@"[A-Z]");
                    var numeric_regex = new Regex(@"[0-9]");
                    var special_regex = new Regex(@"[!@£$%^&*()\[\]:;\\\/<>{}]");
                    var history_regex = new Regex(@"_history[0-9]");

                    CreateSearchTable();
                    foreach (DataRow row in RawDataTable.Rows)
                    {
                        // Set from raw data
                        string username = row[0].ToString();
                        string password = row[1].ToString();

                        string deleetPass = Deleet_password(row[1].ToString());

                        string mask = "";

                        for (int i = 0; i < password.Length; i++)
                        {
                            string c = password[i].ToString();
                            //Process char
                            if ((lower_regex.Match(c).Success))
                            {
                                mask = mask + "?l";
                            }
                            else if ((higher_regex.Match(c).Success))
                            {
                                mask = mask + "?u";
                            }
                            else if ((numeric_regex.Match(c).Success))
                            {
                                mask = mask + "?d";
                            }
                            else if ((special_regex.Match(c).Success))
                            {
                                mask = mask + "?s";
                            }
                        }

                        // Password Check
                        if (RadioSearchPassword.IsChecked == true)
                        {
                            // Variations
                            if (CheckBoxSearchVariation.IsChecked == true)
                            {
                                if (deleetPass.ToLower().Contains(TextBoxSearchInput.Text.ToLower()))
                                {
                                    DataTableSearch.Rows.Add(username, password, mask);
                                }
                            }
                            // Exact
                            else
                            {
                                if (password.Contains(TextBoxSearchInput.Text))
                                {
                                    DataTableSearch.Rows.Add(username, password, mask);
                                }
                            }

                        }

                        // Username Check
                        if (RadioSearchUsername.IsChecked == true)
                        {
                            if (username.ToLower().Contains(TextBoxSearchInput.Text.ToLower()))
                            {
                                DataTableSearch.Rows.Add(username, password, mask);
                            }
                        }
                    }

                    Dispatcher.Invoke(() => DataGridSearchFilter.ItemsSource = DataTableSearch.AsDataView());
                }

            }
            else
            {
                CreateSearchTable();
                Dispatcher.Invoke(() => DataGridSearchFilter.ItemsSource = DataTableSearch.AsDataView());
            }
        }

        private void LoadXMLSummarySettings()
        {
            SummaryXMLVulnAdmins.Text = Properties.Settings.Default.VulnIdAdmins;
            SummaryXMLVulnCharFreq.Text = Properties.Settings.Default.VulnIdCharFreq;
            SummaryXMLVulnCompany.Text = Properties.Settings.Default.VulnIdCompany;
            SummaryXMLVulnDateTime.Text = Properties.Settings.Default.VulnIdDate;
            SummaryXMLVulnDictWords.Text = Properties.Settings.Default.VulnIdDict;
            SummaryXMLVulnHistory.Text = Properties.Settings.Default.VulnIdHistory;
            SummaryXMLVulnKeyboard.Text = Properties.Settings.Default.VulnIdKeyboard;
            SummaryXMLVulnLength.Text = Properties.Settings.Default.VulnIdLength;
            SummaryXMLVulnOverview.Text = Properties.Settings.Default.VulnIdOverview;
            SummaryXMLVulnPwned.Text = Properties.Settings.Default.VulnIdPwned;
            SummaryXMLVulnReuse.Text = Properties.Settings.Default.VulnIdReuse;
            SummaryXMLVulnUsernameInPass.Text = Properties.Settings.Default.VulnIdUsername;
            SummaryXMLVulnWeak.Text = Properties.Settings.Default.VulnIdWeak;
            SummaryXMLSource.Text = Properties.Settings.Default.VulnIdSource;
        }

        private void SaveXMLSumamrySettings()
        {
            Properties.Settings.Default.VulnIdAdmins = SummaryXMLVulnAdmins.Text;
            Properties.Settings.Default.VulnIdCharFreq = SummaryXMLVulnCharFreq.Text;
            Properties.Settings.Default.VulnIdCompany = SummaryXMLVulnCompany.Text;
            Properties.Settings.Default.VulnIdDate = SummaryXMLVulnDateTime.Text;
            Properties.Settings.Default.VulnIdDict = SummaryXMLVulnDictWords.Text;
            Properties.Settings.Default.VulnIdHistory = SummaryXMLVulnHistory.Text;
            Properties.Settings.Default.VulnIdKeyboard = SummaryXMLVulnKeyboard.Text;
            Properties.Settings.Default.VulnIdLength = SummaryXMLVulnLength.Text;
            Properties.Settings.Default.VulnIdOverview = SummaryXMLVulnOverview.Text;
            Properties.Settings.Default.VulnIdPwned = SummaryXMLVulnPwned.Text;
            Properties.Settings.Default.VulnIdReuse = SummaryXMLVulnReuse.Text;
            Properties.Settings.Default.VulnIdUsername = SummaryXMLVulnUsernameInPass.Text;
            Properties.Settings.Default.VulnIdWeak = SummaryXMLVulnWeak.Text;
            Properties.Settings.Default.VulnIdSource = SummaryXMLSource.Text;
            Properties.Settings.Default.Save();
        }

        //// UI Interaction
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void ButtonSettingsDisableAll_Click(object sender, RoutedEventArgs e)
        {
            CheckAnalysisAdministrator.IsChecked = false;
            CheckAnalysisCharacter.IsChecked = false;
            CheckAnalysisCommon.IsChecked = false;
            CheckAnalysisCompany.IsChecked = false;
            CheckAnalysisDate.IsChecked = false;
            CheckAnalysisFrequency.IsChecked = false;
            CheckAnalysisHashcat.IsChecked = false;
            CheckAnalysisHistory.IsChecked = false;
            CheckAnalysisKeyboard.IsChecked = false;
            CheckAnalysisLength.IsChecked = false;
            CheckAnalysisReuse.IsChecked = false;
            CheckAnalysisUsername.IsChecked = false;
            CheckAnalysisPwned.IsChecked = false;
            CheckAnalysisWeak.IsChecked = false;


            TabCharacterAnalysis.Visibility = Visibility.Collapsed;
            TabLengthAnalysis.Visibility = Visibility.Collapsed;
            TabFrequencyAnalysis.Visibility = Visibility.Collapsed;
            TabBaseAnalysis.Visibility = Visibility.Collapsed;
            TabReuseAnalysis.Visibility = Visibility.Collapsed;
            TabKeyboardAnalysis.Visibility = Visibility.Collapsed;
            TabDateAnalysis.Visibility = Visibility.Collapsed;
            TabCompanyAnalysis.Visibility = Visibility.Collapsed;
            TabUsernameAnalysis.Visibility = Visibility.Collapsed;
            TabAdminAnalysis.Visibility = Visibility.Collapsed;
            TabHashcatAnalysis.Visibility = Visibility.Collapsed;
            TabHistoryAnalysis.Visibility = Visibility.Collapsed;
            TabPwnedPasswords.Visibility = Visibility.Collapsed;
            TabWeakAnalysis.Visibility = Visibility.Collapsed;
        }

        private void ButtonAddCompany_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TextBoxCompany.Text))
            {
                if (!ListBoxCompany.Items.Contains(TextBoxCompany.Text))
                {
                    ListBoxCompany.Items.Add(TextBoxCompany.Text);
                    Properties.Settings.Default.CompanyNameListSettings = TextBoxCompany.Text + ";" + Properties.Settings.Default.CompanyNameListSettings;
                }
                Properties.Settings.Default.Save();
            }
        }

        private void ButtonRemoveCompany_Click(object sender, RoutedEventArgs e)
        {
            if (ListBoxCompany.SelectedItems.Count > 0)
            {
                var index = ListBoxCompany.Items.IndexOf(ListBoxCompany.SelectedItem);
                ListBoxCompany.Items.RemoveAt(index);
                Properties.Settings.Default.CompanyNameListSettings = "";
                foreach (string companyName in ListBoxCompany.Items)
                {
                    Properties.Settings.Default.CompanyNameListSettings = companyName + ";" + Properties.Settings.Default.CompanyNameListSettings;
                }

                Properties.Settings.Default.Save();
            }
        }

        private void TextBoxCompany_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBoxCompany.Text = "";
        }

        private void TextBoxCompany_LostFocus(object sender, RoutedEventArgs e)
        {
            if (TextBoxCompany.Text == "")
            {
                TextBoxCompany.Text = "Organisation Name";
            }

        }

        private void ButtonAddAdmin_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TextBoxAdmin.Text))
            {
                if (!ListBoxAdmin.Items.Contains(TextBoxAdmin.Text))
                {
                    ListBoxAdmin.Items.Add(TextBoxAdmin.Text);
                    Properties.Settings.Default.AdminAccountsList = TextBoxAdmin.Text + ";" + Properties.Settings.Default.AdminAccountsList;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private void ButtonRemoveAdmin_Click(object sender, RoutedEventArgs e)
        {
            if (ListBoxAdmin.SelectedItems.Count > 0)
            {
                var index = ListBoxAdmin.Items.IndexOf(ListBoxAdmin.SelectedItem);
                ListBoxAdmin.Items.RemoveAt(index);
                Properties.Settings.Default.AdminAccountsList = "";
                foreach (string admin in ListBoxAdmin.Items)
                {
                    Properties.Settings.Default.AdminAccountsList = admin + ";" + Properties.Settings.Default.AdminAccountsList;
                }

                Properties.Settings.Default.Save();
            }
        }

        private void TextBoxAdmin_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBoxAdmin.Text = "";
        }

        private void TextBoxAdmin_LostFocus(object sender, RoutedEventArgs e)
        {
            if (TextBoxAdmin.Text == "")
            {
                TextBoxAdmin.Text = "Administrator";
            }
        }

        private void ButtonImportAdmin_Click(object sender, RoutedEventArgs e)
        {
            AdminLoadNewFile();
        }

        private void ButtonImportPasswords_Click(object sender, RoutedEventArgs e)
        {
            ImportFilename = ImportUserPassList();
        }

        private void ButtonGenerateSummary_Click(object sender, RoutedEventArgs e)
        {
            GenerateSummary();
        }

        private void ButtonRunAnalysis_Click(object sender, RoutedEventArgs e)
        {
            BoolSummaryAdmin = false;
            BoolSummaryLength = false;
            BoolSummaryCompany = false;
            BoolSummaryCommon = false;
            BoolSummaryUser = false;
            BoolSummaryKeyboard = false;
            BoolSummaryReuse = false;
            BoolSummaryReuseAdmin = false;
            BoolIncludeHistory = false;
            BoolSummaryDate = false;
            BoolSummaryHistory = false;
            BoolSummaryPwned = false;
            BoolSummaryWeak = false;

            bool ContinueAnalysis;
            ClearStatusAndTime();
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            // Delimits file
            // Returns true if successful
            using (new WaitCursor())
            {
                ContinueAnalysis = Dispatcher.Invoke(() => DelimitData(ImportFilename));
            }



            // Runs analysis if delimited file was imported
            if (ContinueAnalysis)
            {
#pragma warning disable CS4014 // Because this call is not awaited, execution of the current method continues before the call is completed
                RunAnalysisChecks();
#pragma warning restore CS4014 // Because this call is not awaited, execution of the current method continues before the call is completed
            }

            // Run to update ADUsers table with Cracked Status
            ADUsersUpdatedCracked();

            stopWatch.Stop();
            TimeSpan ts = stopWatch.Elapsed;

            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);

        }

        // Change filter for password length minimum
        private void ButtonLengthFilter_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisLength.IsChecked.Value)
            {
                try
                {
                    //DataGridLength.ItemsSource = AnalysisLength();
                    Properties.Settings.Default.LengthMinimumUserSetting = TextBoxLengthMinimum.Text.ToString();
                    Properties.Settings.Default.Save();
                }
                catch
                {
                    MessageBox.Show("Please enable Length Analysis to change the filter settings.", "Error: Length Analysis");
                }
            }
        }

        private void DataGridHistoryAll_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {


                var HistoryRow = (DataRowView)DataGridHistoryReuse.SelectedItem;
                string user = HistoryRow[0].ToString();

                DataGridHistorySelect.ItemsSource = HistorySelect(user);
            }
            catch
            {

            }
        }

        // Copy Summary to clipboard
        private void RichTextTech_GotMouseCapture(object sender, MouseEventArgs e)
        {
            string richText = RichTextTech.Text;
            Clipboard.SetText(richText);

        }

        // Copy Summary to clipboard
        private void RichTextExec_GotMouseCapture(object sender, MouseEventArgs e)
        {
            string richText = RichTextExec.Text;
            Clipboard.SetText(richText);
        }

        // Remove all admin usernames from list
        private void ButtonRemoveAdminAll_Click(object sender, RoutedEventArgs e)
        {
            ListBoxAdmin.Items.Clear();

            Properties.Settings.Default.AdminAccountsList = "";
            Properties.Settings.Default.Save();
        }

        // Remove all admin usernames from list
        private void ButtonRemoveCompanyAll_Click(object sender, RoutedEventArgs e)
        {
            ListBoxCompany.Items.Clear();
            Properties.Settings.Default.CompanyNameListSettings = "";
            Properties.Settings.Default.Save();
        }

        private void CreateWordlist_Click(object sender, RoutedEventArgs e)
        {
            string[] WordListRaw;

            WordlistImportFilename = CreateWordList();
            if (WordlistImportFilename.Length > 0)
            {
                WordListRaw = File.ReadAllLines(WordlistImportFilename);
                Task.Run(() => ProcessWordList(WordListRaw));
            }
            else
            {
                MessageBox.Show("Please select a password list or dictionary list to use as a basis for generating a new wordlist.", "Error: Wordlist Builder");
            }
        }

        private void SaveWordlist_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() => SaveWordlist.IsEnabled = false);

            DataTableWordListDistinct.TableName = "WordlistDistinct";

            SaveFileDialog saveWordlistDialog = new SaveFileDialog
            {

                // default save file
                FileName = "pwdlyser-wordlist.txt",
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            };

            if (saveWordlistDialog.ShowDialog() is true)
            {
                Task.Run(() => WriteWordListToFile(saveWordlistDialog.FileName));
            }
            Dispatcher.Invoke(() => SaveWordlist.IsEnabled = true);

        }

        private void CreateWordlistAnalysis_Click(object sender, RoutedEventArgs e)
        {
            string[] RawDataWordListImport = RawDataTable.Rows.OfType<DataRow>().Select(k => k[1].ToString()).ToArray();
            Task.Run(() => ProcessWordList(RawDataWordListImport));
        }

        private void WordlistEnableAll_Click(object sender, RoutedEventArgs e)
        {
            CheckBoxWordlistCamel.IsChecked = true;
            CheckBoxWordlistDeleet.IsChecked = true;
            CheckBoxWordlistLowCase.IsChecked = true;
            CheckBoxWordlistNumbers.IsChecked = true;
            CheckBoxWordlistOriginal.IsChecked = true;
            CheckBoxWordlistSpecial.IsChecked = true;
            CheckBoxWordlistSpecialAll.IsChecked = true;
            CheckBoxWordlistPrefix.IsChecked = true;
            CheckBoxWordlistSuffix.IsChecked = true;
            CheckBoxWordlistFreqSuffix.IsChecked = true;
        }

        private void WordlistDisableAll_Click(object sender, RoutedEventArgs e)
        {
            CheckBoxWordlistCamel.IsChecked = false;
            CheckBoxWordlistDeleet.IsChecked = false;
            CheckBoxWordlistLowCase.IsChecked = false;
            CheckBoxWordlistNumbers.IsChecked = false;
            CheckBoxWordlistOriginal.IsChecked = false;
            CheckBoxWordlistSpecial.IsChecked = false;
            CheckBoxWordlistSpecialAll.IsChecked = false;
            CheckBoxWordlistPrefix.IsChecked = false;
            CheckBoxWordlistSuffix.IsChecked = false;
            CheckBoxWordlistFreqSuffix.IsChecked = false;
        }

        private void ButtonGetStarted_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke((Action)(() => MainTabControl.SelectedIndex = 1));
            Dispatcher.BeginInvoke((Action)(() => AnalysisTabControl.SelectedIndex = 0));
            Properties.Settings.Default.GetStartedAccepted = true;
            Properties.Settings.Default.HideTabs = false;
            Properties.Settings.Default.Save();
            TabItemDashboard.Visibility = Visibility.Collapsed;
            TabAnalysis.Visibility = Visibility.Visible;
            TabWordlist.Visibility = Visibility.Visible;
            TabXMLOutput.Visibility = Visibility.Visible;
        }

        private void ButtonAddBaseword_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TextBoxBaseword.Text))
            {
                if (!ListBoxBaseword.Items.Contains(TextBoxBaseword.Text))
                {
                    ListBoxBaseword.Items.Add(TextBoxBaseword.Text);
                }
            }
        }

        private void ButtonRemoveBaseword_Click(object sender, RoutedEventArgs e)
        {
            if (ListBoxBaseword.SelectedItems.Count > 0)
            {
                var index = ListBoxBaseword.Items.IndexOf(ListBoxBaseword.SelectedItem);
                ListBoxBaseword.Items.RemoveAt(index);
            }
        }

        // Restore Default
        private void ButtonRemoveAllBaseword_Click(object sender, RoutedEventArgs e)
        {


            ListBoxBaseword.Items.Clear();
            string[,] commonList = new string[,]
            {
                {"password"},
                {"123456"},
                {"letmein"},
                {"system"},
                {"admin"},
                {"starwars"},
                {"football"},
                {"welcome"},
                {"abc123"},
                {"dragon"},
                {"iloveyou"},
                {"monkey"},
                {"login"},
                {"soccer"},
                {"freedom"},
                {"love"},
                {"contraseña"}, //Spanish - Password
                {"bienvenido"}, // Spanish - Welcome
                {"déjameentrar"}, // Spanish - Let me in
                {"sistema"}, // Spanish - System
                {"bienvenue"}, // French - Welcome
                {"système"}, // French - system
                {"laissemoientrer"}, // French -Let Me in
                {"welkom"}, // german - welcome
                {"motdepasse"}, // french - password
                {"passwort"}, // Ditch - password
                {"wachtwoord"}, // dutch - password
                {"willkommen" }, // german - welcome
                {"systeem"}, // dutch - system
                {"laatmebinnen" }, // dutch - letmein
                {"lassmichrein" }, // german - letmein
                
            };

            if (ListBoxBaseword.Items.Count == 0)
            {
                foreach (string commonPass in commonList)
                {
                    ListBoxBaseword.Items.Add(commonPass);
                }
            }
        }

        private void TextBoxBaseword_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBoxBaseword.Text = "";
        }

        private void TextBoxBaseword_LostFocus(object sender, RoutedEventArgs e)
        {
            if (TextBoxBaseword.Text == "")
            {
                TextBoxBaseword.Text = "Enter Base-Word";
            }
        }

        // Show hide tabs - Settings
        private void CheckAnalysisCharacter_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisCharacter.IsChecked == true)
            {
                TabCharacterAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabCharacterAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisLength_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisLength.IsChecked == true)
            {
                TabLengthAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabLengthAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisFrequency_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisFrequency.IsChecked == true)
            {
                TabFrequencyAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabFrequencyAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisCommon_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisCommon.IsChecked == true)
            {
                TabBaseAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabBaseAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisReuse_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisReuse.IsChecked == true)
            {
                TabReuseAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabReuseAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisKeyboard_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisKeyboard.IsChecked == true)
            {
                TabKeyboardAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabKeyboardAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisDate_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisDate.IsChecked == true)
            {
                TabDateAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabDateAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisCompany_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisCompany.IsChecked == true)
            {
                TabCompanyAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabCompanyAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisUsername_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisUsername.IsChecked == true)
            {
                TabUsernameAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabUsernameAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisAdministrator_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisAdministrator.IsChecked == true)
            {
                TabAdminAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabAdminAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisHashcat_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisHashcat.IsChecked == true)
            {
                TabHashcatAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabHashcatAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void CheckAnalysisHistory_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisHistory.IsChecked == true)
            {
                TabHistoryAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabHistoryAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        // Manually Add Total over time
        private void ButtonRunTotal_Click(object sender, RoutedEventArgs e)
        {

            AddToProjectXML();

        }

        private void TotalSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var filePath = Environment.ExpandEnvironmentVariables(Properties.Settings.Default.TotalOverTimeSettingsPath);
            string dirPath = Path.GetDirectoryName(filePath);

            OpenFileDialog dlg = new OpenFileDialog
            {
                DefaultExt = ".xml",
                InitialDirectory = dirPath
            };

            bool? result = dlg.ShowDialog();

            if (result.HasValue && result.Value)
            {
                Properties.Settings.Default.TotalOverTimeSettingsPath = dlg.FileName;
                Properties.Settings.Default.Save();
                LoadTotalXMLFile();
            }
        }

        private void TotalAddProject_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveNewProject = new SaveFileDialog
            {
                FileName = "",
                Filter = "Pwdlyser Project File|*.xml|All files (*.*)|*.*"
            };

            string filePath = "";
            if (saveNewProject.ShowDialog() is true)
            {
                filePath = saveNewProject.FileName;
            }

            using (StreamWriter sw = File.CreateText(filePath))
            {
                sw.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sw.WriteLine("<DocumentElement>");
                sw.WriteLine("<Item date=\"1970-01-01\" total =\"0\"/>");
                sw.WriteLine("</DocumentElement>");

            }
            Properties.Settings.Default.TotalOverTimeSettingsPath = filePath;
            Properties.Settings.Default.Save();

            LoadTotalXMLFile();
        }

        private void ChartTotalCracked_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (OverviewHiddenTab.Visibility == Visibility.Visible)
            {
                OverviewHiddenTab.Visibility = Visibility.Collapsed;
            }
            else if (OverviewHiddenTab.Visibility == Visibility.Collapsed)
            {
                OverviewHiddenTab.Visibility = Visibility.Visible;
            }
        }

        private void ChartFrequency_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowFrequency.Visibility == Visibility.Visible)
            {
                HideShowFrequency.Visibility = Visibility.Collapsed;
            }
            else if (HideShowFrequency.Visibility == Visibility.Collapsed)
            {
                HideShowFrequency.Visibility = Visibility.Visible;
            }
        }

        private void ChartLength_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowLength.Visibility == Visibility.Visible)
            {
                HideShowLength.Visibility = Visibility.Collapsed;
            }
            else if (HideShowLength.Visibility == Visibility.Collapsed)
            {
                HideShowLength.Visibility = Visibility.Visible;
            }
        }

        private void ChartCharacter_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowCharacter.Visibility == Visibility.Visible)
            {
                HideShowCharacter.Visibility = Visibility.Collapsed;
            }
            else if (HideShowCharacter.Visibility == Visibility.Collapsed)
            {
                HideShowCharacter.Visibility = Visibility.Visible;
            }
        }

        private void ChartCommon_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowBaseword.Visibility == Visibility.Visible)
            {
                HideShowBaseword.Visibility = Visibility.Collapsed;
            }
            else if (HideShowBaseword.Visibility == Visibility.Collapsed)
            {
                HideShowBaseword.Visibility = Visibility.Visible;
            }
        }

        private void ChartReuse_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowReuse.Visibility == Visibility.Visible)
            {
                HideShowReuse.Visibility = Visibility.Collapsed;
            }
            else if (HideShowReuse.Visibility == Visibility.Collapsed)
            {
                HideShowReuse.Visibility = Visibility.Visible;
            }
        }

        private void ChartUsername_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowUsername.Visibility == Visibility.Visible)
            {
                HideShowUsername.Visibility = Visibility.Collapsed;
            }
            else if (HideShowUsername.Visibility == Visibility.Collapsed)
            {
                HideShowUsername.Visibility = Visibility.Visible;
            }
        }

        private void ChartAdmin_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowAdmins.Visibility == Visibility.Visible)
            {
                HideShowAdmins.Visibility = Visibility.Collapsed;
            }
            else if (HideShowAdmins.Visibility == Visibility.Collapsed)
            {
                HideShowAdmins.Visibility = Visibility.Visible;
            }
        }

        private void ChartCompany_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowCompany.Visibility == Visibility.Visible)
            {
                HideShowCompany.Visibility = Visibility.Collapsed;
            }
            else if (HideShowCompany.Visibility == Visibility.Collapsed)
            {
                HideShowCompany.Visibility = Visibility.Visible;
            }
        }

        private void ChartDate_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowDate.Visibility == Visibility.Visible)
            {
                HideShowDate.Visibility = Visibility.Collapsed;
            }
            else if (HideShowDate.Visibility == Visibility.Collapsed)
            {
                HideShowDate.Visibility = Visibility.Visible;
            }
        }

        private void ChartKeyboard_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowKeyboard.Visibility == Visibility.Visible)
            {
                HideShowKeyboard.Visibility = Visibility.Collapsed;
            }
            else if (HideShowKeyboard.Visibility == Visibility.Collapsed)
            {
                HideShowKeyboard.Visibility = Visibility.Visible;
            }
        }

        private void CheckAnalysisPwned_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisPwned.IsChecked == true)
            {
                TabPwnedPasswords.Visibility = Visibility.Visible;
            }
            else
            {
                TabPwnedPasswords.Visibility = Visibility.Collapsed;
            }
        }

        private void ChartPwned_MouseDown(object sender, MouseButtonEventArgs e)
        {

            if (HideShowPwned.Visibility == Visibility.Visible)
            {
                HideShowPwned.Visibility = Visibility.Collapsed;
            }
            else if (HideShowPwned.Visibility == Visibility.Collapsed)
            {
                HideShowPwned.Visibility = Visibility.Visible;
            }
        }

        private void ButtonSearchDomain_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() => RetrieveUserDetailsForAnalysis());
        }

        private void CheckADEnabled_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckADEnabled.IsChecked == true && DataGridUserProperties.Items.Count > 0)
            {
                ADEnableFilter();

            }
        }

        private void CheckADDisabled_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckADDisabled.IsChecked == true && DataGridUserProperties.Items.Count > 0)
            {
                ADDisabledFilter();
            }
        }

        private void CheckADAll_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckADAll.IsChecked == true && DataGridUserProperties.Items.Count > 0)
            {
                Dispatcher.Invoke(() => DataGridUserProperties.ItemsSource = ADUsers.AsDataView());

            }
        }

        private void CheckEnableADPassword_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckEnableADPassword.IsChecked == true)
            {
                ADGridCredentials.Visibility = Visibility.Visible;
            }
            else
            {
                ADGridCredentials.Visibility = Visibility.Collapsed;
            }
        }

        private void ChartWeak_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (HideShowWeak.Visibility == Visibility.Visible)
            {
                HideShowWeak.Visibility = Visibility.Collapsed;
            }
            else if (HideShowWeak.Visibility == Visibility.Collapsed)
            {
                HideShowWeak.Visibility = Visibility.Visible;
            }
        }

        private void CheckAnalysisWeak_Click(object sender, RoutedEventArgs e)
        {
            if (CheckAnalysisWeak.IsChecked == true)
            {
                TabWeakAnalysis.Visibility = Visibility.Visible;
            }
            else
            {
                TabWeakAnalysis.Visibility = Visibility.Collapsed;
            }
        }

        private void ButtonSaveDomainFile_Click(object sender, RoutedEventArgs e)
        {
            SaveDomainData();
        }

        private void ButtonLoadDomainFile_Click(object sender, RoutedEventArgs e)
        {
            LoadDomainData();
        }

        // Dynamic search
        private void TextBoxSearchInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            DynamicSearch();
        }

        private void RadioSearchUsername_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxSearchInput.Text = "";

            try
            {
                if (CheckBoxSearchVariation.IsChecked == true)
                {
                    CheckBoxSearchVariation.IsChecked = false;
                }
            }

            catch
            {

            }
        }

        private void RadioSearchPassword_Checked(object sender, RoutedEventArgs e)
        {
            TextBoxSearchInput.Text = "";
        }

        private void CheckBoxSearchVariation_Checked(object sender, RoutedEventArgs e)
        {
            DynamicSearch();
        }

        private void CheckBoxSearchVariation_Unchecked(object sender, RoutedEventArgs e)
        {
            DynamicSearch();
        }

        private void CheckADCracked_Click(object sender, RoutedEventArgs e)
        {
            if (CheckADCracked.IsChecked == true && DataGridUserProperties.Items.Count > 0)
            {

                ADCrackedFilter();
            }
        }

        private void CheckBoxBaseWordDict_Checked(object sender, RoutedEventArgs e)
        {
            if (CheckBoxBaseWordDict.IsChecked == true)
            {
                ListBoxBaseword.IsEnabled = false;
                TextBoxBaseword.IsEnabled = false;
                ButtonAddBaseword.IsEnabled = false;
                ButtonRemoveBaseword.IsEnabled = false;
                ButtonRemoveAllBaseword.IsEnabled = false;
                Dispatcher.Invoke(() => CheckBoxBaseWordDictLarge.IsChecked = false);
            }
        }

        private void CheckBoxBaseWordDictLarge_Checked(object sender, RoutedEventArgs e)
        {

            if (CheckBoxBaseWordDictLarge.IsChecked == true)
            {
                ListBoxBaseword.IsEnabled = false;
                TextBoxBaseword.IsEnabled = false;
                ButtonAddBaseword.IsEnabled = false;
                ButtonRemoveBaseword.IsEnabled = false;
                ButtonRemoveAllBaseword.IsEnabled = false;
                Dispatcher.Invoke(() => CheckBoxBaseWordDict.IsChecked = false);
            }
        }

        private void CheckBoxBaseWordDict_Unchecked(object sender, RoutedEventArgs e)
        {

            if ((CheckBoxBaseWordDictLarge.IsChecked == false) && (CheckBoxBaseWordDictLarge.IsChecked == false))
            {
                ListBoxBaseword.IsEnabled = true;
                TextBoxBaseword.IsEnabled = true;
                ButtonAddBaseword.IsEnabled = true;
                ButtonRemoveBaseword.IsEnabled = true;
                ButtonRemoveAllBaseword.IsEnabled = true;
                Dispatcher.Invoke(() => CheckBoxBaseWordDict.IsChecked = false);
                Dispatcher.Invoke(() => CheckBoxBaseWordDictLarge.IsChecked = false);
            }
        }

        private void ButtonSummarySaveXML_Click(object sender, RoutedEventArgs e)
        {
            // write to user properties //summaryxml
            Dispatcher.Invoke(() => SaveXMLSumamrySettings());
        }

        private void ButtonSummarySaveXMLFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveXmlDlg = new SaveFileDialog();

            saveXmlDlg.FileName = "pwdlyser-xml-" + SummaryXMLDomain.Text + ".xml";
            saveXmlDlg.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            var filePath = saveXmlDlg.FileName;

            // set a default file name
            // set filters - this can be done in properties as well

            if (saveXmlDlg.ShowDialog() == true)
            {
                using (StreamWriter sw = new StreamWriter(saveXmlDlg.FileName))
                {
                    sw.WriteLine((FormatXml(StringSummaryXMLOutput).Replace("hostname=\"\"", "hostname=\"" + SummaryXMLDomain.Text + "\"")));
                }
            }
        }

        private void ButtonSummaryRegenXMLFile_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() => RichTextXML.Text = "");
            SummaryOutputTechnical();
        }
    }
}
