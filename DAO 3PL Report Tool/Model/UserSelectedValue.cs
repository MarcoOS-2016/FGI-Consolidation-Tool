using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAO_3PL_Report_Tool
{
    public class UserSelectedValue
    {
        private DateTime begindate;
        private DateTime enddate;
        private DateTime snapshotdate;
        private string sourcefolder;
        private string outputfolder;
        
        private string agingreportfolder;
        private string detailsreportfolder;
        private string onhandreportfolder;
        private string agingreportoutputfolder;

        private string weeknumber;
        private string regions;
        private string lobs;
        private string categories;

        private string apj3plreportsourcefolder;
        private string apjsnppartlistfile;
        private string apjoutputfolder;

        public DateTime BeginDate
        {
            get { return begindate; }
            set { begindate = value; }
        }

        public DateTime EndDate
        {
            get { return enddate; }
            set { enddate = value; }
        }

        public DateTime SnapshotDate
        {
            get { return snapshotdate; }
            set { snapshotdate = value; }
        }

        public string SourceFolder
        {
            get { return sourcefolder; }
            set { sourcefolder = value; }
        }

        public string OutputFolder
        {
            get { return outputfolder; }
            set { outputfolder = value; }
        }

        public string WeekNumber
        {
            get { return weeknumber; }
            set { weeknumber = value; }
        }

        public string AgingReportFolder
        {
            get { return agingreportfolder; }
            set { agingreportfolder = value; }
        }

        public string DetailsReportFolder
        {
            get { return detailsreportfolder; }
            set { detailsreportfolder = value; }
        }

        public string OnhandReportFolder
        {
            get { return onhandreportfolder; }
            set { onhandreportfolder = value; }
        }

        public string AgingReportOutputFolder
        {
            get { return agingreportoutputfolder; }
            set { agingreportoutputfolder = value; }
        }

        public string Regions
        {
            get { return regions; }
            set { regions = value; }
        }

        public string LOBs
        {
            get { return lobs; }
            set { lobs = value; }
        }

        public string Catogories
        {
            get { return categories; }
            set { categories = value; }
        }

        public string APJ3PLReportSourceFolder
        {
            get { return apj3plreportsourcefolder; }
            set { apj3plreportsourcefolder = value; }
        }

        public string APJSnPPartListFile
        {
            get { return apjsnppartlistfile; }
            set { apjsnppartlistfile = value; }
        }

        public string APJOutputFolder
        {
            get { return apjoutputfolder; }
            set { apjoutputfolder = value; }
        }
    }
}
