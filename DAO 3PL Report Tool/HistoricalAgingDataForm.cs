using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DAO_3PL_Report_Tool
{
    public partial class HistoricalAgingDataForm : Form
    {
        private DataTable ConsolidationAgingDataTable = null;
        private UserSelectedValue userSelectedValue = null;

        public HistoricalAgingDataForm()
        {
            InitializeComponent();
        }

        private void HistoricalAgingDataForm_Load(object sender, EventArgs e)
        {
            ExtensionMethods.DoubleBuffered(historicalAgingDataGridView, true);

            List<string> weekNumberList = new List<string>();

            string beginWeekNumber = ConfigFileUtility.GetValue("BeginWeekNumber");
            string endWeekNumber = ConfigFileUtility.GetValue("EndWeekNumber");
            string[] regions = ConfigFileUtility.GetValue("RegionList").ToUpper().Split(',');
            string[] lobs = ConfigFileUtility.GetValue("LOBList").ToUpper().Split(',');
            string[] categories = ConfigFileUtility.GetValue("CategoryList").ToUpper().Split(',');

            string prefixWeekNumber = beginWeekNumber.Substring(0, 6);
            int beginNumber = 
                Convert.ToInt32(beginWeekNumber.Substring(prefixWeekNumber.Length, beginWeekNumber.Length - prefixWeekNumber.Length));
            int endNumber =
                Convert.ToInt32(endWeekNumber.Substring(prefixWeekNumber.Length, endWeekNumber.Length - prefixWeekNumber.Length));

            for (int index = beginNumber; index <= endNumber; index++)
            {
                weekNumberList.Add(string.Format("{0}{1}", prefixWeekNumber, index.ToString()));
            }

            //hist_aging_WeekNumberComboBox.DataSource = weekNumberList;
            WeekNumberCheckBoxComboBox.DataSource = MainForm.WeekNumberList;

            foreach (string region in regions)
                regionCheckBoxComboBox.Items.Add(region);

            foreach (string lob in lobs)
                lobCheckBoxComboBox.Items.Add(lob);

            foreach (string category in categories)
                categoryCheckBoxComboBox.Items.Add(category);
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            if (historicalAgingDataGridView.RowCount == 0)
            {
                MessageBox.Show("No data can be exported!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                SaveFileDialog savefile = new SaveFileDialog();
                savefile.Filter = "CSV file (*.csv)|All files (*.*)";
                savefile.FilterIndex = 1;
                savefile.AddExtension = true;
                savefile.RestoreDirectory = true;

                if (savefile.ShowDialog() == DialogResult.OK)
                {
                    string filename = string.Format("{0}.csv", savefile.FileName);

                    if (historicalAgingDataGridView.RowCount != 0)
                    {
                        DataTable datatable = FillDataTable("historicalAgingDataGridView");

                        ExcelFileUtility.ExportDataIntoCSVFile(filename, datatable);
                        MessageBox.Show("Report has been generated successful!",
                                "Export Data",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    string.Format("Function name:[Export Aging Report] - Message:{0}, Source:{1}, StackTrack:{2}", ex.Message, ex.Source, ex.StackTrace));
                throw ex;
            }
        }

        private void queryButton_Click(object sender, EventArgs e)
        {
            userSelectedValue = new UserSelectedValue();
            userSelectedValue.WeekNumber = WeekNumberCheckBoxComboBox.SelectedItem.ToString();

            if (regionCheckBoxComboBox.Text.Length != 0)
                userSelectedValue.Regions = regionCheckBoxComboBox.Text.ToUpper().Trim();

            if (lobCheckBoxComboBox.Text.Length != 0)
                userSelectedValue.LOBs = lobCheckBoxComboBox.Text.ToUpper().Trim();

            if (categoryCheckBoxComboBox.Text.Length != 0)
                userSelectedValue.Catogories = categoryCheckBoxComboBox.Text.ToUpper().Trim();

            if (ConsolidationAgingDataTable == null)
                StartSynchronizedJob("LoadConsolidationAgingReport");

            ShowConsolidationAgingData();
        }

        private void ShowConsolidationAgingData()
        {
            // Filter raw data by conditions which are choiced by user
            if (userSelectedValue.Regions != null || userSelectedValue.LOBs != null || userSelectedValue.Catogories != null)
            {
                DataTable datatable = FilterRawData();
                historicalAgingDataGridView.DataSource = datatable;
            }
            else
                historicalAgingDataGridView.DataSource = ConsolidationAgingDataTable;
        }

        private void LoadConsolidationAgingReport()
        {
            string consolidationAgingReportOutputFolder = ConfigFileUtility.GetValue("ConsolidationAgingReportOutputFolder");
            string consolidationAgingFilePrefix = ConfigFileUtility.GetValue("ConsolidationAgingFilePrefix");

            string fileName = string.Format("{0}_{1}.csv", consolidationAgingFilePrefix, userSelectedValue.WeekNumber);
            string fullFileName = Path.Combine(consolidationAgingReportOutputFolder, fileName);

            if (!File.Exists(fullFileName))
            {
                MessageBox.Show(string.Format("Cannot find {0} report file, Please check if it is existed in the folder - {1}.",
                    fullFileName, consolidationAgingReportOutputFolder),
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                MiscUtility.LogHistory(string.Format("Starting to load Consolidation Aging report - {0}...", fullFileName));
                //MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to load Consolidation Aging report - {1}...",
                //    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), filename));

                string textLine = string.Empty;
                List<string[]> lineList = new List<string[]>();

                StreamReader fileReader = new StreamReader(fullFileName);
                textLine = fileReader.ReadLine();
                string[] titleLineText = textLine.Split('\t');

                BuildConsolidationAgingDataTableStructure(titleLineText);

                while (textLine != null)
                {
                    textLine = fileReader.ReadLine();
                    if (textLine != null && textLine.Length > 0)
                    {
                        FillConsolidationAgingDataTable(textLine);
                    }
                }

                MiscUtility.LogHistory("Done!");
                //MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        private DataTable FilterRawData()
        {
            StringBuilder regions = new StringBuilder();
            StringBuilder lobs = new StringBuilder();
            StringBuilder categories = new StringBuilder();

            if (userSelectedValue.Regions != null)
            {
                foreach (string region in userSelectedValue.Regions.Split(','))
                    regions.Append(string.Format("'{0}',", region.Trim()));
                regions.Remove(regions.Length - 1, 1);
            }
            else
                regions.Append("'NULL'");

            if (userSelectedValue.LOBs != null)
            {
                foreach (string lob in userSelectedValue.LOBs.Split(','))
                    lobs.Append(string.Format("'{0}',", lob.Trim()));
                lobs.Remove(lobs.Length - 1, 1);
            }
            else
                lobs.Append("'NULL'");

            if (userSelectedValue.Catogories != null)
            {
                foreach (string category in userSelectedValue.Catogories.Split(','))
                    categories.Append(string.Format("'{0}',", category.Trim()));
                categories.Remove(categories.Length - 1, 1);
            }
            else
                categories.Append("'NULL'");

            string expression = 
                string.Format("(REGION in ({0}) or 'NULL' in ({0})) and (LOB in ({1}) or 'NULL' in ({1})) and (REVISED_CATEGORY in ({2}) or 'NULL' in ({2}))", 
                regions.ToString(), lobs.ToString(), categories.ToString());

            DataTable datatable = ConsolidationAgingDataTable.Clone();
            datatable.Clear();
            
            DataRow[] foundRows = ConsolidationAgingDataTable.Select(expression);
            for (int index = 0; index < foundRows.Length; index++)
            {
                datatable.Rows.Add(foundRows[index].ItemArray);
            }

            return datatable;
        }

        private void ExportDataGrid()
        {

        }

        private void BuildConsolidationAgingDataTableStructure(string[] titlelinetext)
        {
            ConsolidationAgingDataTable = new DataTable();
            DataColumn dc = null;

            for (int index = 0; index < titlelinetext.Length - 1; index++)
            {
                dc = new DataColumn();
                dc.ColumnName = titlelinetext[index].ToUpper();
                ConsolidationAgingDataTable.Columns.Add(dc);
            }
        }

        private void FillConsolidationAgingDataTable(string textLine)
        {
            string[] textCell = textLine.Split('\t');
            DataRow dataRow = ConsolidationAgingDataTable.NewRow();
            
            for (int index = 0; index < textCell.Length - 1; index++)
            {                
                dataRow[index] = textCell[index];
            }

            ConsolidationAgingDataTable.Rows.Add(dataRow);
        }

        private DataTable FillDataTable(string controlname)
        {
            DataTable dt = new DataTable();

            DataColumn dc = null;
            DataRow dr = null;

            DataGridView datagridview = (DataGridView)this.Controls.Find(controlname, true)[0];

            for (int headcount = 0; headcount < datagridview.ColumnCount; headcount++)
            {
                dc = new DataColumn();
                dc.ColumnName = datagridview.Columns[headcount].HeaderText;
                dt.Columns.Add(dc);
            }

            for (int rowcount = 0; rowcount < datagridview.RowCount - 1; rowcount++)
            {
                dr = dt.NewRow();
                for (int colcount = 0; colcount < datagridview.ColumnCount; colcount++)
                {
                    dr[colcount] = datagridview[colcount, rowcount].Value;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        #region ----- Backgroundworker -----
        ProgessForm progressform = null;
        public static BackgroundWorker backgroundworker = new BackgroundWorker();

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
                case "LoadConsolidationAgingReport":
                    LoadConsolidationAgingReport();
                    break;
            }
        }

        public void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //toolStripStatusLabel.Text = e.UserState.ToString();
        }

        public void CompletedWork(object sender, RunWorkerCompletedEventArgs e)
        {
            if (progressform != null)
            {
                progressform.Hide();
                progressform = null;

                //MessageBox.Show("The processing has been completed successful!", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
    }
}
