using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace DAO_3PL_Report_Tool
{
    public class DAO3PLReportFileHandler
    {
        private readonly UserSelectedValue UserSelectedValue = null;
        private List<ThirdPartyFile> unprocessedfilenamelist = null;
        private List<string> tranfieldvaluelist = new List<string>();
        private List<string> snapfieldvaluelist = new List<string>();        
        private DataTable onhanddatatable = null;
        private DataTable snapshotdatatable = null;
        private DataTable transactiondatatable = null;

        public DAO3PLReportFileHandler(UserSelectedValue userSelectValue)
        {
            this.UserSelectedValue = userSelectValue;
        }

        public void Process()
        {
            GetUnProcessedFileNameList();
            BuildDataTableStructure();
            ReadFileContent();
            BuildSnapShotDataTable();
            ExportReport();
        }

        private void GetUnProcessedFileNameList()
        {
            //bool flag = false;
            DirectoryInfo dir = null;
            string fileName = string.Empty;
            string textLine = string.Empty;
            string tempFileName = string.Empty;
            string tranFilePostfixNameList = ConfigFileUtility.GetValue("TranFilePostfixNameList");
            string snapFilePostfixNameList = ConfigFileUtility.GetValue("SnapFilePostfixNameList");
            
            try
            {                
                MiscUtility.LogHistory("Starting to search unprocessed DAO 3PL report files...");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to search unprocessed DAO 3PL report files...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

                unprocessedfilenamelist = new List<ThirdPartyFile>();
                dir = new DirectoryInfo(UserSelectedValue.SourceFolder);

                foreach (FileInfo fi in dir.GetFiles())
                {
                    fileName = fi.Name;
                    
                    if (fileName.ToUpper().Contains("TRAN_3PL") && fi.Extension.ToLower().Equals(".txt"))
                    {
                        foreach (string filePostfixNameList in tranFilePostfixNameList.Split(','))
                        {
                            if (fileName.Contains(filePostfixNameList))
                            {

                                if ((fi.LastWriteTime.Date.CompareTo(UserSelectedValue.BeginDate) == 0 || fi.LastWriteTime.Date.CompareTo(UserSelectedValue.BeginDate) > 0)
                                    && (fi.LastWriteTime.Date.CompareTo(UserSelectedValue.EndDate) == 0 || fi.LastWriteTime.Date.CompareTo(UserSelectedValue.EndDate) < 0))
                                {
                                    ThirdPartyFile thirdPartyTextFile = new ThirdPartyFile();
                                    thirdPartyTextFile.FileName = fi.Name;
                                    thirdPartyTextFile.FullFillName = fi.FullName;
                                    thirdPartyTextFile.FileExtension = fi.Extension;

                                    unprocessedfilenamelist.Add(thirdPartyTextFile);
                                }
                            }
                        }
                    }

                    if (fileName.ToUpper().Contains("SNAP_3PL") && fi.Extension.ToLower().Equals(".txt"))
                    {
                        foreach (string filePostfixNameList in snapFilePostfixNameList.Split(','))
                        {
                            if (fileName.Contains(filePostfixNameList))
                            {
                                if (fi.LastWriteTime.Date.CompareTo(UserSelectedValue.SnapshotDate) == 0)
                                {
                                    ThirdPartyFile thirdPartyTextFile = new ThirdPartyFile();
                                    thirdPartyTextFile.FileName = fi.Name;
                                    thirdPartyTextFile.FullFillName = fi.FullName;
                                    thirdPartyTextFile.FileExtension = fi.Extension;

                                    unprocessedfilenamelist.Add(thirdPartyTextFile);
                                }
                            }
                        }
                    }                        
                }

                MiscUtility.LogHistory(string.Format("Total of {0} files found!", unprocessedfilenamelist.Count));
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Total of {1} files found!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), unprocessedfilenamelist.Count));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <GetUnProcessedFileNameList>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void ReadFileContent()
        {
            string textLine = string.Empty;
            string message = string.Empty;

            try
            {
                foreach (ThirdPartyFile thirdPartyFile in this.unprocessedfilenamelist)
                {
                    using (StreamReader reader = new FileInfo(thirdPartyFile.FullFillName).OpenText())
                    {
                        if (thirdPartyFile.FileName.ToUpper().Contains("TRAN_3PL"))
                        {
                            MiscUtility.LogHistory(string.Format("Starting to read 3PL Tran text file - {0}...", thirdPartyFile.FullFillName));
                            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to read 3PL Tran text file - {1}...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), thirdPartyFile.FullFillName));

                            tranfieldvaluelist.Clear();
                            textLine = reader.ReadLine();

                            while ((textLine = reader.ReadLine()) != null)
                            {
                                if (textLine.ToUpper().Contains("RECT") && !textLine.ToUpper().Contains("SHIP") && !textLine.ToUpper().Contains("ADJS"))            // Filter text line with prefix "RECT"
                                {
                                    foreach (string tempLine in textLine.Split('|'))
                                    {
                                        tranfieldvaluelist.Add(tempLine.Replace("\"", ""));
                                    }
                                }
                            }

                            reader.Close();
                            FillDataTable("TRAN");
                            
                            MiscUtility.LogHistory("Done!");
                            MainForm.backgroundworker.ReportProgress(0, "Done!");
                        }

                        if (thirdPartyFile.FileName.ToUpper().Contains("SNAP_3PL"))
                        {
                            message = string.Format("Starting to read 3PL Snapshot text file - {0}...", thirdPartyFile.FullFillName);
                            
                            MiscUtility.LogHistory(message);
                            MainForm.backgroundworker.ReportProgress(0, message);

                            snapfieldvaluelist.Clear();
                            textLine = reader.ReadLine();

                            while ((textLine = reader.ReadLine()) != null)
                            {
                                foreach (string tempLine in textLine.Split('|'))
                                {
                                    snapfieldvaluelist.Add(tempLine.Replace("\"", ""));
                                }
                            }

                            reader.Close();
                            FillDataTable("SNAP");
                                                        
                            MiscUtility.LogHistory("Done!");
                            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                        }                    
                    }
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <ReadFileContent>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void FillDataTable(string identifier)
        {
            DataRow dataRow = null;

            if (identifier.Equals("TRAN"))
            {
                for (int index = 0; index < tranfieldvaluelist.Count; index += 12)
                {
                    dataRow = transactiondatatable.NewRow();

                    dataRow["Record_Type"] = tranfieldvaluelist[index].ToString().Trim();
                    dataRow["Date_Time"] = tranfieldvaluelist[index + 1].ToString().Trim();
                    dataRow["From_Site_ID"] = tranfieldvaluelist[index + 2].ToString().Trim();
                    dataRow["To_Site_ID"] = tranfieldvaluelist[index + 3].ToString().Trim();
                    dataRow["Transaction_Type"] = tranfieldvaluelist[index + 4].ToString().Trim();
                    dataRow["Part"] = tranfieldvaluelist[index + 5].ToString().Trim();
                    dataRow["Qty"] = tranfieldvaluelist[index + 6].ToString().Trim();
                    dataRow["Unique Ref"] = tranfieldvaluelist[index + 7].ToString().Trim();
                    dataRow["Transaction ID"] = tranfieldvaluelist[index + 8].ToString().Trim();
                    dataRow["Alt_Ref"] = tranfieldvaluelist[index + 9].ToString().Trim();
                    dataRow["Service_Tag"] = tranfieldvaluelist[index + 10].ToString().Trim();
                    dataRow["Part_Class"] = tranfieldvaluelist[index + 11].ToString().Trim();

                    transactiondatatable.Rows.Add(dataRow);
                }
            }

            if (identifier.Equals("SNAP"))
            {
                for (int index = 0; index < snapfieldvaluelist.Count; index += 6)
                {
                    dataRow = onhanddatatable.NewRow();
                    
                    dataRow["Site_ID"] = snapfieldvaluelist[index + 1].ToString().Trim();
                    dataRow["Part"] = snapfieldvaluelist[index + 2].ToString().Trim();

                    if (snapfieldvaluelist[index + 3].ToString().Trim().Length == 0)
                        dataRow["Qty"] = 0;
                    else
                        dataRow["Qty"] = snapfieldvaluelist[index + 3].ToString().Trim();

                    onhanddatatable.Rows.Add(dataRow);
                }
            }
        }

        private void BuildSnapShotDataTable()
        {
            bool flag = false;
            int quantity = 0;
            string thirdName = string.Empty;
            string partNumber = string.Empty;
            DataRow dataRow = null;
                        
            MiscUtility.LogHistory("Starting to build Snapshot datatable...");
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to build Snapshot datatable...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));

            for (int index = 0; index < onhanddatatable.Rows.Count; index++)
            {
                flag = false;                
                thirdName = onhanddatatable.Rows[index][0].ToString();
                partNumber = onhanddatatable.Rows[index][1].ToString();
                quantity = Convert.ToInt32(onhanddatatable.Rows[index][2]);

                for (int indey = 0; indey < snapshotdatatable.Rows.Count; indey ++)
                {
                    if (partNumber.Equals(snapshotdatatable.Rows[indey][0].ToString()))
                    {
                        switch (thirdName)
                        {
                            case "3PL EG3":
                                snapshotdatatable.Rows[indey][1] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][1]) + quantity;
                                snapshotdatatable.Rows[indey][8] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][8]) + quantity;
                                break;

                            case "AW1":
                                snapshotdatatable.Rows[indey][2] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][2]) + quantity;
                                snapshotdatatable.Rows[indey][8] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][8]) + quantity;
                                break;

                            case "BNA":
                                snapshotdatatable.Rows[indey][3] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][3]) + quantity;
                                snapshotdatatable.Rows[indey][8] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][8]) + quantity;
                                break;

                            case "ELP":
                                snapshotdatatable.Rows[indey][4] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][4]) + quantity;
                                snapshotdatatable.Rows[indey][8] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][8]) + quantity;
                                break;

                            case "FG3":
                                snapshotdatatable.Rows[indey][5] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][5]) + quantity;
                                snapshotdatatable.Rows[indey][8] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][8]) + quantity;
                                break;

                            case "MXC":
                                snapshotdatatable.Rows[indey][6] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][6]) + quantity;
                                snapshotdatatable.Rows[indey][8] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][8]) + quantity;
                                break;

                            case "SCA":
                                snapshotdatatable.Rows[indey][7] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][7]) + quantity;
                                snapshotdatatable.Rows[indey][8] =
                                    Convert.ToInt32(snapshotdatatable.Rows[indey][8]) + quantity;
                                break;
                        }

                        flag = true;
                        break;
                    }
                }

                if (!flag)
                {
                    dataRow = snapshotdatatable.NewRow();

                    switch (thirdName)
                    {
                        case "3PL EG3":
                            dataRow["FGAID"] = partNumber;
                            dataRow["3PL EG3"] = quantity;
                            dataRow["AW1"] = 0;
                            dataRow["BNA"] = 0;
                            dataRow["ELP"] = 0;
                            dataRow["FG3"] = 0;
                            dataRow["MXC"] = 0;
                            dataRow["SCA"] = 0;
                            dataRow["Total_Qty"] = quantity;
                            break;

                        case "AW1":
                            dataRow["FGAID"] = partNumber;
                            dataRow["3PL EG3"] = 0;
                            dataRow["AW1"] = quantity;
                            dataRow["BNA"] = 0;
                            dataRow["ELP"] = 0;
                            dataRow["FG3"] = 0;
                            dataRow["MXC"] = 0;
                            dataRow["SCA"] = 0;
                            dataRow["Total_Qty"] = quantity;
                            break;

                        case "BNA":
                            dataRow["FGAID"] = partNumber;
                            dataRow["3PL EG3"] = 0;
                            dataRow["AW1"] = 0;
                            dataRow["BNA"] = quantity;
                            dataRow["ELP"] = 0;
                            dataRow["FG3"] = 0;
                            dataRow["MXC"] = 0;
                            dataRow["SCA"] = 0;
                            dataRow["Total_Qty"] = quantity;
                            break;

                        case "ELP":
                            dataRow["FGAID"] = partNumber;
                            dataRow["3PL EG3"] = 0;
                            dataRow["AW1"] = 0;
                            dataRow["BNA"] = 0;
                            dataRow["ELP"] = quantity;
                            dataRow["FG3"] = 0;
                            dataRow["MXC"] = 0;
                            dataRow["SCA"] = 0;
                            dataRow["Total_Qty"] = quantity;
                            break;

                        case "FG3":
                            dataRow["FGAID"] = partNumber;
                            dataRow["3PL EG3"] = 0;
                            dataRow["AW1"] = 0;
                            dataRow["BNA"] = 0;
                            dataRow["ELP"] = 0;
                            dataRow["FG3"] = quantity;
                            dataRow["MXC"] = 0;
                            dataRow["SCA"] = 0;
                            dataRow["Total_Qty"] = quantity;
                            break;

                        case "MXC":
                            dataRow["FGAID"] = partNumber;
                            dataRow["3PL EG3"] = 0;
                            dataRow["AW1"] = 0;
                            dataRow["BNA"] = 0;
                            dataRow["ELP"] = 0;
                            dataRow["FG3"] = 0;
                            dataRow["MXC"] = quantity;
                            dataRow["SCA"] = 0;
                            dataRow["Total_Qty"] = quantity;
                            break;

                        case "SCA":
                            dataRow["FGAID"] = partNumber;
                            dataRow["3PL EG3"] = 0;
                            dataRow["AW1"] = 0;
                            dataRow["BNA"] = 0;
                            dataRow["ELP"] = 0;
                            dataRow["FG3"] = 0;
                            dataRow["MXC"] = 0;
                            dataRow["SCA"] = quantity;
                            dataRow["Total_Qty"] = quantity;
                            break;
                    }

                    snapshotdatatable.Rows.Add(dataRow);                    
                }
            }

            MiscUtility.LogHistory("Done!");
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
        }

        private void ExportReport()
        {
            string consolidatedReceiptFileName = Path.Combine(UserSelectedValue.OutputFolder,
                string.Format("{0}_{1}.csv", "Consolidated_Receipt", DateTime.Now.ToString("yyyy-MM-dd_HHmm")));
            string onhandFileName = Path.Combine(UserSelectedValue.OutputFolder,
                string.Format("{0}_{1}.csv", "OnHand_Inventory", DateTime.Now.ToString("yyyy-MM-dd_HHmm")));

            ExportToExcelFile(consolidatedReceiptFileName, transactiondatatable);
            ExportToExcelFile(onhandFileName, snapshotdatatable);
        }

        private void ExportToExcelFile(string filename, DataTable datatable)
        {                        
            MiscUtility.LogHistory(string.Format("Starting to export data into Excel file - {0}...", filename));
            MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Starting to export data into Excel file - {1}...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), filename));

            try
            {
                ExcelFileUtility.ExportDataIntoCSVFile(filename, datatable);
                //MiscUtility.SaveFile(thirdFileName, filename);

                MiscUtility.LogHistory("Done!");
                MainForm.backgroundworker.ReportProgress(0, string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <ExportToExcelFile>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        #region --- Build table structure ---
        private void BuildDataTableStructure()
        {
            BuildTransactionTableStructure();
            BuildOnHandDataTableStructure();
            BuildSnapShotDataTableStructure();
        }
        
        private void BuildTransactionTableStructure()
        {
            transactiondatatable = new DataTable();

            DataColumn dc = new DataColumn();
            dc.ColumnName = "Record_Type";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Date_Time";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "From_Site_ID";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "To_Site_ID";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Transaction_Type";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Part";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Qty";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Unique Ref";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Transaction ID";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Alt_Ref";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Service_Tag";
            transactiondatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Part_Class";
            transactiondatatable.Columns.Add(dc);
        }

        private void BuildOnHandDataTableStructure()
        {
            onhanddatatable = new DataTable();

            DataColumn dc = new DataColumn();
            dc.ColumnName = "Site_ID";
            onhanddatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Part";
            onhanddatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Qty";
            onhanddatatable.Columns.Add(dc);
        }

        private void BuildSnapShotDataTableStructure()
        {
            snapshotdatatable = new DataTable();

            DataColumn dc = new DataColumn();
            dc.ColumnName = "FGAID";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "3PL EG3";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "AW1";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "BNA";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "ELP";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "FG3";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "MXC";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "SCA";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Total_Qty";
            snapshotdatatable.Columns.Add(dc);
        }
        #endregion
    }
}
