using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace DMLSS_Attachment_Handler
{
    public abstract class Version
    {
        //check if current version of this appliation is the same as updated version.
        public static Boolean checkVersion()
        {
            if (getUpdatedVersion() == getCurrentSWVersion())
                return true;
            else
            {
                updateRequired();
                return false;
            }
        }

        //getUpdatedVersion method is to check a file in network drive and return version.
        public static string getUpdatedVersion()
        {
            string version = System.IO.File.ReadAllText(@"N:\DMLSS\Version\DMLSS_AH");
            version = version.Trim();
            version = version.Substring(version.LastIndexOf("=") + 1);
            return version;
        }

        //getCurrentSWVersion method is to retrun current version of this application.
        public static string getCurrentSWVersion()
        {
            return DMLSS_Attachment.version;
        }

        //updateRequired() method is to desplay and warn the version is outdated.
        public static void updateRequired()
        {
            MessageBox.Show("Application version is NOT correct. Please update the application." + "\nCurrent Version: "+getCurrentSWVersion()+"\nNew Version: "+getUpdatedVersion());
        }
    }
}
