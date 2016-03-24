using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Resources;
using System.Threading;
using System.Windows;
using System.Windows.Forms;

namespace DAO_3PL_Report_Tool
{
    public partial class MainForm : Form
    {
        public static List<string> WeekNumberList = BuildWeekNumberList();
        public static BackgroundWorker backgroundworker = new BackgroundWorker();
        private UserSelectedValue userselectedvalue = new UserSelectedValue();

        public MainForm()
        {
            InitializeComponent();
            SetDefaultValue();
        }

        private void SetDefaultValue()
        {
            // ---- Set default values for Tab page "Consolidate APJ 3PL Report" ----
            apj3PLReport_Tab_SourceReportTextBox.Text = ConfigFileUtility.GetValue("APJ3PLSourceReportFolder");
            apj3PLReport_Tab_SnPPartListFileTextBox.Text = ConfigFileUtility.GetValue("APJSnPPartListFile");
            apj3PLReport_Tab_OutputFolderTextBox.Text = ConfigFileUtility.GetValue("APJConsolidationReportOutputFolder");

            // ---- Set default values for Tab page "Consolidate DAO 3PL Report" ----
            dao3PLReport_Tab_sourceFolderTextBox.Text = ConfigFileUtility.GetValue("DAO3PLReportSourceFolder");
            dao3PLReport_Tab_outputFolderTextBox.Text = ConfigFileUtility.GetValue("DAO3PLConsolidationReportOutputFolder");

            // ---- Set default values for Tab page "Generate Aging Report" ----
            generateAging_Tab_WeekNumberComboBox.DataSource = WeekNumberList;
            generateAging_Tab_AgingReportFolderTextBox.Text = ConfigFileUtility.GetValue("AgingReportSourceFolder");
            generateAging_Tab_DetailsReportFolderTextBox.Text = ConfigFileUtility.GetValue("DetailsReportSourceFolder");
            generateAging_Tab_OnhandReportTextBox.Text = ConfigFileUtility.GetValue("OnhandReportSourceFolder");
            generateAging_Tab_InTransitReportTextBox.Text = ConfigFileUtility.GetValue("InTransitReportSourceFolder");
            generateAging_Tab_OutputFolderTextBox.Text = ConfigFileUtility.GetValue("ConsolidationAgingReportOutputFolder");
        }

        #region ----- Tab page - DAO 3PL Report Consolidation -----
        private void sourceFolderButton_Click(object sender, EventArgs e)
        {
            dao3PLReport_Tab_sourceFolderTextBox.Text = SelectFolder();
        }

        private void outputFolderButton_Click(object sender, EventArgs e)
        {
            dao3PLReport_Tab_outputFolderTextBox.Text = SelectFolder();
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            DateTime beginDate = dao3PLReport_Tab_beginDateTimePicker.Value.Date;
            DateTime endDate = dao3PLReport_Tab_endDateTimePicker.Value.Date;
            DateTime snapshotDate = dao3PLReport_Tab_snapShotDateTimePicker.Value.Date;

            string sourceFolder = dao3PLReport_Tab_sourceFolderTextBox.Text.Trim();
            string outputFolder = dao3PLReport_Tab_outputFolderTextBox.Text.Trim();

            toolStripStatusLabel.Text = "";

            if (dao3PLReport_Tab_beginDateTimePicker.Value.Date.CompareTo(DateTime.Now.Date) == 0)
            {
                if (MessageBox.Show("Do you want to select current day as Begin Date?", "Confirm", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
                return;
            }

            if (dao3PLReport_Tab_endDateTimePicker.Value.Date.CompareTo(DateTime.Now.Date) == 0)
            {
                if (MessageBox.Show("Do you want to select current day as End Date?", "Confirm", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
                    return;
            }

            if (dao3PLReport_Tab_snapShotDateTimePicker.Value.Date.CompareTo(DateTime.Now.Date) == 0)
            {
                if (MessageBox.Show("Do you want to select current day as Snapshot Date?", "Confirm", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
                    return;
            }

            if (dao3PLReport_Tab_beginDateTimePicker.Value.Date.CompareTo(DateTime.Now.Date) > 0)
            {
                MessageBox.Show("The Begin Date is later than today!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (dao3PLReport_Tab_beginDateTimePicker.Value.Date.CompareTo(dao3PLReport_Tab_endDateTimePicker.Value.Date) > 0)
            {
                MessageBox.Show("The Begin Date is later than End Date!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (dao3PLReport_Tab_endDateTimePicker.Value.Date.CompareTo(DateTime.Now.Date) > 0)
            {
                MessageBox.Show("The End Date is later than today!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (dao3PLReport_Tab_endDateTimePicker.Value.Date.CompareTo(dao3PLReport_Tab_beginDateTimePicker.Value.Date) < 0)
            {
                MessageBox.Show("The End Date is earlier than Begin Date!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (dao3PLReport_Tab_snapShotDateTimePicker.Value.Date.CompareTo(DateTime.Now.Date) > 0)
            {
                MessageBox.Show("The Snapshot Date is later than today!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (sourceFolder.Length == 0)
            {
                MessageBox.Show("Please select a source folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (outputFolder.Length == 0)
            {
                MessageBox.Show("Please select an output folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            userselectedvalue.BeginDate = beginDate;
            userselectedvalue.EndDate = endDate;
            userselectedvalue.SnapshotDate = snapshotDate;
            userselectedvalue.SourceFolder = sourceFolder;
            userselectedvalue.OutputFolder = outputFolder;

            StartSynchronizedJob("Consolidate3PLReports");

            toolStripStatusLabel.Text = "Completed!";
        }
        #endregion

        #region ----- Tab page - Generate Aging Report -----
        private void generateAging_Tab_AgingReportOpenButton_Click(object sender, EventArgs e)
        {
            generateAging_Tab_AgingReportFolderTextBox.Text = SelectFolder();
        }
        
        private void generateAging_Tab_DetailsReportOpenButton_Click(object sender, EventArgs e)
        {
            generateAging_Tab_DetailsReportFolderTextBox.Text = SelectFolder();
        }

        private void generateAging_Tab_OnhandReportOpenButton_Click(object sender, EventArgs e)
        {
            generateAging_Tab_OnhandReportTextBox.Text = SelectFolder();
        }

        private void generateAging_Tab_InTransitReportButton_Click(object sender, EventArgs e)
        {
            generateAging_Tab_InTransitReportTextBox.Text = SelectFolder();
        } 

        private void generateAging_Tab_OutputFolderOpenButton_Click(object sender, EventArgs e)
        {
            generateAging_Tab_OutputFolderTextBox.Text = SelectFolder();
        }

        private void generateAging_Tab_startButton_Click(object sender, EventArgs e)
        {
            string agingReportFolder = generateAging_Tab_AgingReportFolderTextBox.Text.Trim();
            string detailsReportFolder = generateAging_Tab_DetailsReportFolderTextBox.Text.Trim();
            string onhandReportFolder = generateAging_Tab_OnhandReportTextBox.Text.Trim();
            string intransitReportFolder = generateAging_Tab_InTransitReportTextBox.Text.Trim();
            string agingreportoutputFolder = generateAging_Tab_OutputFolderTextBox.Text.Trim();

            if (agingReportFolder.Length == 0)
            {
                MessageBox.Show("Please select an Aging Report Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (detailsReportFolder.Length == 0)
            {
                MessageBox.Show("Please select a Details Report Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (onhandReportFolder.Length == 0)
            {
                MessageBox.Show("Please select a On-hand Report Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (intransitReportFolder.Length == 0)
            {
                MessageBox.Show("Please select a In-Transit Report Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (agingreportoutputFolder.Length == 0)
            {
                MessageBox.Show("Please select a Output Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            userselectedvalue.WeekNumber = generateAging_Tab_WeekNumberComboBox.SelectedValue.ToString();
            userselectedvalue.AgingReportFolder = agingReportFolder;
            userselectedvalue.DetailsReportFolder = detailsReportFolder;
            userselectedvalue.OnhandReportFolder = onhandReportFolder;
            userselectedvalue.AgingReportOutputFolder = agingreportoutputFolder;

            StartSynchronizedJob("BuildAgingReport");
        }
        #endregion

        #region ----- Tab page - APJ 3PL Report Consolidation -----
        private void apj3PLReport_Tab_SourceReportButton_Click(object sender, EventArgs e)
        {
            apj3PLReport_Tab_SourceReportTextBox.Text = SelectFolder();
        }

        private void apj3PLReport_Tab_SnPPartListFileButton_Click(object sender, EventArgs e)
        {
            apj3PLReport_Tab_SnPPartListFileTextBox.Text = SelectFile();
        }

        private void apj3PLReport_Tab_OutputFolderButton_Click(object sender, EventArgs e)
        {
            apj3PLReport_Tab_OutputFolderTextBox.Text = SelectFolder();
        }

        private void apj3PLReport_Tab_StartButton_Click(object sender, EventArgs e)
        {
            string sourceReportFolder = apj3PLReport_Tab_SourceReportTextBox.Text.Trim();
            string snpPartListFile = apj3PLReport_Tab_SnPPartListFileTextBox.Text.Trim();
            string outputFolder = apj3PLReport_Tab_OutputFolderTextBox.Text.Trim();

            if (sourceReportFolder.Length == 0)
            {
                MessageBox.Show("Please select an APJ Source Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (snpPartListFile.Length == 0)
            {
                MessageBox.Show("Please select a SnP Part List File", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (outputFolder.Length == 0)
            {
                MessageBox.Show("Please select an Output Folder", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            UserSelectedValue userSelectedValue = new UserSelectedValue();
            userselectedvalue.APJ3PLReportSourceFolder = sourceReportFolder;
            userselectedvalue.APJSnPPartListFile = snpPartListFile;
            userselectedvalue.APJOutputFolder = outputFolder;

            StartSynchronizedJob("ConsolidateAPJ3PLReports");
        }
        #endregion

        private void ConsolidateAPJ3PLReports()
        {
            APJ3PLReporFileHandler handler = new APJ3PLReporFileHandler(userselectedvalue);
            handler.Process();
        }

        private void ConsolidateDAO3PLReports()
        {
            DAO3PLReportFileHandler handler = new DAO3PLReportFileHandler(userselectedvalue);
            handler.Process();
        }

        private void BuildAgingReport()
        {
            AgingReportHandler handler = new AgingReportHandler(userselectedvalue);
            handler.Process();
        }

        #region ----- Backgroundworker -----        
        ProgessForm progressform = null;

        public void StartSynchronizedJob(object instance)
        {
            backgroundworker.WorkerReportsProgress = true;
            backgroundworker.WorkerSupportsCancellation = true;
            backgroundworker.DoWork += new DoWorkEventHandler(DoWork);
            backgroundworker.ProgressChanged += new ProgressChangedEventHandler(ProgressChanged);
            backgroundworker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(CompletedWork);

            progressform = new ProgessForm();
            backgroundworker.RunWorkerAsync(instance);

            progressform.ShowDialog(this);
            progressform = null;
        }

        public void DoWork(object sender, DoWorkEventArgs e)
        {
            string functionName = (string)e.Argument;

            switch (functionName)
            {
                case "ConsolidateAPJ3PLReports":
                    ConsolidateAPJ3PLReports();
                    break;

                case "ConsolidateDAO3PLReports":
                    ConsolidateDAO3PLReports();
                    break;

                case "BuildAgingReport":
                    BuildAgingReport();
                    break;
            }
        }

        public void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripStatusLabel.Text = e.UserState.ToString();
        }

        public void CompletedWork(object sender, RunWorkerCompletedEventArgs e)
        {
            if (progressform != null)
            {
                progressform.Hide();                
                progressform = null;

                MessageBox.Show("The processing has been completed successful!", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            // Check to see if an error occured in the 
            // background process.
            if (e.Error != null)
            {
                //IsError = true;
                MessageBox.Show(e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //MessageBox.Show(e.Error.Message);
                return;
            }

            // Check to see if the background process was cancelled.
            if (e.Cancelled)
            {
                MessageBox.Show("Processing cancelled!", "Cancel", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        #endregion

        #region ----- Common functions -----
        private string SelectFolder()
        {
            FolderBrowserDialog folderbrowser = new FolderBrowserDialog();
            folderbrowser.RootFolder = Environment.SpecialFolder.MyComputer;
            folderbrowser.SelectedPath = @"C:\";
            folderbrowser.ShowNewFolderButton = true;

            if (folderbrowser.ShowDialog() == DialogResult.OK)
            {
                return folderbrowser.SelectedPath;
            }

            return String.Empty;
        }

        private string SelectFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = @"C:\";
            openFileDialog.Filter = "Excel file (*.xls, *.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }

            return String.Empty;
        }
        #endregion

        private void historicalAgingDataMenuItem_Click(object sender, EventArgs e)
        {
            HistoricalAgingDataForm historicalAgingForm = new HistoricalAgingDataForm();
            historicalAgingForm.ShowDialog();
        }

        public static List<string> BuildWeekNumberList()
        {
            List<string> weekNumberList = new List<string>();

            string beginWeekNumber = ConfigFileUtility.GetValue("BeginWeekNumber");
            string endWeekNumber = ConfigFileUtility.GetValue("EndWeekNumber");

            string prefixWeekNumber = beginWeekNumber.Substring(0, 6);
            int beginNumber =
                Convert.ToInt32(beginWeekNumber.Substring(prefixWeekNumber.Length, beginWeekNumber.Length - prefixWeekNumber.Length));
            int endNumber =
                Convert.ToInt32(endWeekNumber.Substring(prefixWeekNumber.Length, endWeekNumber.Length - prefixWeekNumber.Length));

            for (int index = beginNumber; index <= endNumber; index++)
            {
                weekNumberList.Add(string.Format("{0}{1}", prefixWeekNumber, index.ToString()));
            }

            return weekNumberList;
        }      
    }

    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }
}
