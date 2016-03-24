using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DAO_3PL_Report_Tool
{
    public class ThirdPartyFile
    {
        private string filename;
        private string fullfilename;
        private string fileextension;
        private string site;

        public string FileName
        {
            get { return filename; }
            set { filename = value; }
        }

        public string FullFillName
        {
            get { return fullfilename; }
            set { fullfilename = value; }
        }

        public string FileExtension
        {
            get { return fileextension; }
            set { fileextension = value; }
        }

        public string Site
        {
            get { return site; }
            set { site = value; }
        }
    }
}
