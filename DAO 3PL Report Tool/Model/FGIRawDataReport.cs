using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAO_3PL_Report_Tool
{
    public class FGIOriginalReport
    {
        private string fgidetailsreportfile;
        private string apjagingreportfile;
        private string apjonhandreportfile;
        private string apjintransitreportfile;
        private string daoagingreportfile;
        private string daoonhandreportfile;
        private string daointransitreportfile;
        private string emeaagingreportfile;
        private string emeaonhandreportfile;
        private string emeaintransitreportfile;

        public string FGIDetailsReportFile
        {
            set { fgidetailsreportfile = value; }
            get { return fgidetailsreportfile; }
        }

        public string APJAgingReportFile
        {
            set { apjagingreportfile = value; }
            get { return apjagingreportfile; }
        }

        public string APJOnhandReportFile
        {
            set { apjonhandreportfile = value; }
            get { return apjonhandreportfile; }
        }

        public string APJInTransitReportFile
        {
            set { apjintransitreportfile = value; }
            get { return apjintransitreportfile; }
        }

        public string DAOAgingReportFile
        {
            set { daoagingreportfile = value; }
            get { return daoagingreportfile; }
        }

        public string DAOOnhandReportFile
        {
            set { daoonhandreportfile = value; }
            get { return daoonhandreportfile; }
        }

        public string DAOInTransitReportFile
        {
            set { daointransitreportfile = value; }
            get { return daointransitreportfile; }
        }

        public string EMEAAgingReportFile
        {
            set { emeaagingreportfile = value; }
            get { return emeaagingreportfile; }
        }

        public string EMEAOnhandReportFile
        {
            set { emeaonhandreportfile = value; }
            get { return emeaonhandreportfile; }
        }

        public string EMEAInTransitReportFile
        {
            set { emeaintransitreportfile = value; }
            get { return emeaintransitreportfile; }
        }
    }
}
