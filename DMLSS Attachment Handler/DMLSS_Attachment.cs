using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace DMLSS_Attachment_Handler
{
    //Dialog Answers
    public enum DIAL_ANS { YES, NO, CANCEL }
    class DMLSS_Attachment
    {
        //version number
        public static readonly string version = "1.0";

        private string[] MONTH = { "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC" };
        bool isTypeFilled = false;
        bool isWOFilled = false;
        bool isECNFilled = false;
        bool isSPRFilled = false;
        bool isOTHERFilled = true;

        private string nameOfOther = "";
        private string sourceFilePath = "";
        private string destinationFilePath = "";

        //Check if N drive is currently mapped and connected.
        public bool checkNetworkDrive()
        {
            if(System.IO.Directory.Exists(@"N:\"))
            {
                return true;
            }
            else
            {
                MessageBox.Show("Please check if your network drive is mapped.");
            }
            return false;
        }

        public void setTypeFilled(bool check)
        {
            isTypeFilled = check;
        }
        public void setWOFilled(bool check)
        {
            isWOFilled = check;
        }
        public void setECNFilled(bool check)
        {
            isECNFilled = check;
        }
        public void setSPRFilled(bool check)
        {
            isSPRFilled = check;
        }

        public void setOTHERFilled(bool check)
        {
            isOTHERFilled = check;
        }
    
        //getFileName method is take workodernumber, ecn, and type, change the name format for the name convention.
        public string getFileName(string workorderNumber, string ECN, string type)
        {
            string fileName = "";
            if (ECN.Length == 0)
            {
                fileName = workorderNumber + "_" + type;
            } 
            else if (ECN.Length <= 6)
            {
                for (int i = ECN.Length; i < 6; i++)
                {
                    ECN = "0" + ECN;
                }
                fileName = workorderNumber + "_" + ECN + "_" + type;
            }
            return fileName;
        }

        public void setSourcePath(string path)
        {
            sourceFilePath = path;
        }

        public string getSourcePath()
        {
            return sourceFilePath;
        }

        public void setDestinationPath(string path)
        {
            destinationFilePath = path;
        }

        //setDestinationPath is to set the destination folder with changed file name which is following the name convention.
        public void setDestinationPath(string workorderNumber, string ECN, string type, int month)
        {
            string yearText = getYearFromWO(workorderNumber);
            string monthText = getMonthFromWO(workorderNumber);
            if (type=="SPR")
            {
                int currentMonth = DateTime.Now.Month-1;
                //int sprMonth = Int32.Parse(workorderNumber.Substring(4, 2));
                int sprMonth = month;
                if (currentMonth > sprMonth)
                {
                    int currentYear = DateTime.Now.Year;
                    yearText = (currentYear + 1).ToString();
                }
                monthText = (sprMonth+1).ToString()+"_"+MONTH[sprMonth];
                destinationFilePath = @"N:\DMLSS\" + type + "\\" + yearText + "\\" + monthText + "\\" + getFileName(workorderNumber, ECN, type) + ".pdf";
            }
            else if (type == "Others")
            {
                destinationFilePath = @"N:\DMLSS\" + type + "\\" + yearText + "\\" + getFileName(workorderNumber, ECN, getOtherName()) + ".pdf";
            }
            else
            {
                destinationFilePath = @"N:\DMLSS\" + type + "\\" + yearText + "\\" + getFileName(workorderNumber, ECN, type) + ".pdf";
            }
        }

        public void setOtherName(string name)
        {
            nameOfOther = name;
        }

        public string getOtherName()
        {
            return nameOfOther;
        }

        public string getDestinationPath()
        {
            return destinationFilePath;
        }

        //checkDestinationFIle method to check if the file exists and not empty.
        public bool checkDestinationFile()
        {
            string file = getDestinationPath();
            if(File.Exists(file))
            {
                long fileSize = new System.IO.FileInfo(file).Length;
                if(fileSize > 1)
                {
                    return true;
                }
            }
            return false;
        }

        //checkDestinationDirectory is to check the folder exists where the file will save in.
        public bool checkDestinationDirectory()
        {
            if (System.IO.Directory.GetParent(getDestinationPath()).Exists)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private string getYearFromWO(string wo)
        {
            return wo.Substring(0, 4);
        }

        public string getMonthFromWO(string wo)
        {
            int index = Int32.Parse(wo.Substring(4, 2));
            return MONTH[index-1];
        }

        //checkMonth is to validate the month.
        public bool checkMonth(string wo)
        {
            int index = Int32.Parse(wo.Substring(4, 2));
            if (index > 12)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        //isReadyToUpload method is to check all the required filled before uploade.
        public bool isReadyToUpload()
        {
            if(isTypeFilled && isWOFilled && isECNFilled && isSPRFilled && isOTHERFilled)
            {
                return true;
            }
            return false;
        }

        //uploadFile method is to validate if there is any duplications, and upload the workorder file.
        public bool uploadFile()
        {
            if(checkDestinationDirectory())
            {
                if (File.Exists(getDestinationPath()))
                {
                    DIAL_ANS answer = dialogBox("Over-Writing", "File already exists, do you want to overwrite?", MessageBoxButton.YesNoCancel);
                    if (answer == DIAL_ANS.YES)
                    {
                        File.Copy(getSourcePath(), getDestinationPath(), true);
                        return true;
                    }
                    else if (answer == DIAL_ANS.NO)
                    {
                        int nextNumber = 1;
                        string newDestinationPath = getDestinationPath();
                        newDestinationPath = newDestinationPath.Substring(0, newDestinationPath.LastIndexOf("."));
                        while (File.Exists(newDestinationPath+".pdf"))
                        {
                            if (newDestinationPath.EndsWith(")"))
                            {
                                newDestinationPath = newDestinationPath.Substring(0, newDestinationPath.LastIndexOf("("));
                            }
                            newDestinationPath = newDestinationPath + "(" +nextNumber++ + ")";
                        }
                        File.Copy(getSourcePath(), newDestinationPath+".pdf", false);
                        setDestinationPath(newDestinationPath);
                        return true;
                    }
                }
                else
                {
                    File.Copy(getSourcePath(), getDestinationPath(), false);
                    return true;
                }
            }
            else
            {
                MessageBox.Show("Destination Folder does not exist.");
            }
            return false;
        }

        //dialogBox method to take parameters to display a dialog box and return the answer.
        public DIAL_ANS dialogBox(string title, string message, MessageBoxButton mbb)
        {
            MessageBoxResult answer = MessageBox.Show(message, title, mbb);
            switch(answer)
            {
                case MessageBoxResult.Yes:
                    return DIAL_ANS.YES;
                case MessageBoxResult.No:
                    return DIAL_ANS.NO;
                case MessageBoxResult.Cancel:
                    return DIAL_ANS.CANCEL;
            }
            return DIAL_ANS.CANCEL;
        }

        //debugDisplay is to display string with messagebox.
        public void debugDisplay(string msg)
        {
            MessageBox.Show(msg);
        }

        //debugDisplay is to display integer with messagebox.
        public void debugDisplay(int msg)
        {
            MessageBox.Show(msg.ToString());
        }
    }
}
