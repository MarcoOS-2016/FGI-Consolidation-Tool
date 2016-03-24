using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace DAO_3PL_Report_Tool
{
    public class AppSetting
    {
        private string key;
        private string keyvalue;

        public string Key
        {
            get { return key; }
            set { key = value; }
        }

        public string KeyValue
        {
            get { return keyvalue; }
            set { keyvalue = value; }
        }
    }

    public class ConfigFileUtility
    {
        private static Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

        public static List<AppSetting> GetKeyValueList()
        {
            List<AppSetting> keyvaluelist = new List<AppSetting>();
            KeyValueConfigurationCollection settings = config.AppSettings.Settings;

            foreach (KeyValueConfigurationElement keyValueElement in settings)
            {
                AppSetting keyvalue = new AppSetting();
                keyvalue.Key = keyValueElement.Key.Trim();
                keyvalue.KeyValue = keyValueElement.Value.Trim();

                keyvaluelist.Add(keyvalue);
            }

            return keyvaluelist;
        }

        public static string GetValue(string key)
        {
            string strReturn = null;
            if (config.AppSettings.Settings[key] != null)
            {
                strReturn = config.AppSettings.Settings[key].Value.Trim();
            }

            return strReturn;
        }

        public static void SetValue(string key, string value)
        {
            if (config.AppSettings.Settings[key] != null)
            {
                config.AppSettings.Settings[key].Value = value;
            }
            else
            {
                config.AppSettings.Settings.Add(key, value.Trim());
            }

            config.Save(ConfigurationSaveMode.Modified);
        }

        public static void DelValue(string key)
        {
            config.AppSettings.Settings.Remove(key);
        }
    }
}
