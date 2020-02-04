using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.OutlookViewCtl;
using Path = System.IO.Path;

enum TYPE_OF_DOC { GPC = 0, QUOTE, SR, INVOICE, ECAT, SPR, OTHER }
namespace DMLSS_Attachment_Handler
{  
    //This is the main window
    public partial class MainWindow : Window
    {
        private readonly Microsoft.Office.Interop.Outlook.Application _application = new Microsoft.Office.Interop.Outlook.Application();
        DMLSS_Attachment da = new DMLSS_Attachment();
        string typeOfDoc = "";
        public MainWindow()
        {
            if (Version.checkVersion())
            {
                InitializeComponent();
                setDefault();
                attachmentPath.TextWrapping = TextWrapping.NoWrap;
            }
            else
            {
                Environment.Exit(1);
            }
        }

        //Completely close the application.
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        //setDefault method is to initialize the default setting of fields.
        private void setDefault()
        {
            TypeOfDoc.IsEnabled = false;
            WON_TEXTBOX.IsEnabled = false;
            ECN_TEXTBOX.IsEnabled = false;
            LabelOfSPR.Visibility = Visibility.Hidden;
            MonthOfSPR.Visibility = Visibility.Hidden;
            LabelOfOTHER.Visibility = Visibility.Hidden;
            NameOfOTHER.Visibility = Visibility.Hidden;
            attachmentPath.Text="";
            attachmentPath.IsReadOnly = true;
        }

        //TypeOfDoc_SelectionChanged event handler to handle any change on type of document.
        private void TypeOfDoc_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (TypeOfDoc.SelectedIndex)
            {
                case (int)TYPE_OF_DOC.GPC:
                    typeOfDoc = "GPC";
                    break;
                case (int)TYPE_OF_DOC.QUOTE:
                    typeOfDoc = "Quote";
                    break;
                case (int)TYPE_OF_DOC.SR:
                    typeOfDoc = "SR";
                    break;
                case (int)TYPE_OF_DOC.INVOICE:
                    typeOfDoc = "Invoice";
                    break;
                case (int)TYPE_OF_DOC.ECAT:
                    typeOfDoc = "ECAT";
                    break;
                case (int)TYPE_OF_DOC.SPR:
                    typeOfDoc = "SPR";
                    LabelOfSPR.Visibility = Visibility.Visible;
                    MonthOfSPR.Visibility = Visibility.Visible;
                    break;
                case (int)TYPE_OF_DOC.OTHER:
                    typeOfDoc = "Others";
                    LabelOfOTHER.Visibility = Visibility.Visible;
                    NameOfOTHER.Visibility = Visibility.Visible;
                    da.setOTHERFilled(false);
                    break;
            }
            if (TypeOfDoc.SelectedIndex!=5) 
            {

                LabelOfSPR.Visibility = Visibility.Hidden;
                MonthOfSPR.Visibility = Visibility.Hidden;
               
            } 
            else if (TypeOfDoc.SelectedIndex != 6)
            {
                LabelOfOTHER.Visibility = Visibility.Hidden;
                NameOfOTHER.Visibility = Visibility.Hidden;
                da.setOTHERFilled(true);
            }

            WON_TEXTBOX.IsEnabled = true;
            ECN_TEXTBOX.IsEnabled = true;
        }

        private int GetTextboxLength(TextBox tb)
        {
            return tb.Text.Length;
        }

        //WON_TEXTBOX_Filter event handler to validate if the field has only numeric value.
        private void WON_TEXTBOX_Filter(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }

        //WON_TEXTBOX_TextChanged event handler to validate if the work order field is 12 digit.
        //If the length of work order number is not 12 digit, border line is red. otherwise, it is black.
        private void WON_TEXTBOX_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (GetTextboxLength(WON_TEXTBOX) != 12)
            {
                WON_TEXTBOX.BorderBrush = System.Windows.Media.Brushes.Red;
            }
            else
            {
                WON_TEXTBOX.BorderBrush = System.Windows.Media.Brushes.Black;
            }
        }

        private void enableTypeBox()
        {
            TypeOfDoc.IsEnabled = true;
        }

        //Clear_Click event handler to reset all the fields.
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            reset();
        }

        //reset method to set the program as default
        private void reset()
        {
            pdfWebViewer.Navigate("about:blank");
            da = new DMLSS_Attachment();
            WON_TEXTBOX.Text = "";
            ECN_TEXTBOX.Text = "";
            TypeOfDoc.SelectedIndex = -1;
            MonthOfSPR.SelectedIndex = -1;
            UPLOAD_BUTTON.IsEnabled = true;
            setDefault();
        }

        //pdfWebViewer_Navigated event handler is triggered whenever pdfWebViewer navigated.
        private void pdfWebViewer_Navigated(object sender, NavigationEventArgs e)
        {
            string file = e.Uri.AbsolutePath;
            file = file.Replace("%20",@" ");
            if (File.Exists(file))
            {
                da.setSourcePath(file);
                TypeOfDoc.IsEnabled = true;
            }
        }

        //validateFields is to validate fields.
        private void validateFields()
        {
            if (TypeOfDoc.SelectedIndex != -1)
            {
                da.setTypeFilled(true);
            }
            else
            {
                da.setTypeFilled(false);
            }
            if (TypeOfDoc.SelectedIndex != 5)
            {
                da.setSPRFilled(true);
            }
            else if(MonthOfSPR.SelectedIndex!=-1)
            {
                da.setSPRFilled(true);
            }
            else
            {
                da.setSPRFilled(false);
            }
            if (WON_TEXTBOX.Text.Length == 12)
            {
                da.setWOFilled(true);
            }
            else
            {
                da.setWOFilled(false);
            }
            if (NameOfOTHER.Text.Length!=0)
            {
                da.setOTHERFilled(true);
            }
        }

        //Upload_Click event handler for upload files.
        private void Upload_Click(object sender, RoutedEventArgs e)
        {
            if (da.checkMonth(WON_TEXTBOX.Text))
            {
                if (NameOfOTHER.Text.Length > 0)
                {
                    da.setOtherName(NameOfOTHER.Text);
                }
                da.setDestinationPath(WON_TEXTBOX.Text, ECN_TEXTBOX.Text, typeOfDoc, MonthOfSPR.SelectedIndex);
                validateFields();

                if (da.checkDestinationDirectory())
                {
                    if (ECN_TEXTBOX.Text.Length == 0)
                    {
                        DIAL_ANS answer = da.dialogBox("No ECN", "You didn't enter ECN, Do you want to continue?", MessageBoxButton.YesNo);
                        if (answer == DIAL_ANS.YES)
                        {
                            da.setECNFilled(true);
                        }
                        else
                        {
                            da.setECNFilled(false);
                        }
                    }
                    else
                    {
                        da.setECNFilled(true);
                    }
                    if (da.isReadyToUpload())
                    {
                        if (da.uploadFile())
                        {
                            attachmentPath.Text = da.getDestinationPath();
                            MessageBox.Show("Your file is successfully uploaded");
                            UPLOAD_BUTTON.IsEnabled = false;
                        }
                        else
                        {
                            MessageBox.Show("Upload has been cancelled");
                        }
                        
                    }
                    else
                    {
                        MessageBox.Show("Please complete all required fields");
                    }
                }
                else
                {
                    MessageBox.Show("Please check Work Order Number or Destination Folder.");
                }
            }
            else
            {
                MessageBox.Show("Please check Work Order Number.");
            }
        }

        private void MonthOfSPR_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(MonthOfSPR.SelectedIndex!=-1)
            {
                da.setSPRFilled(true);
            }
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(da.getDestinationPath());
            MessageBox.Show("Copied, you can paste it now.");
        }

        private string FixFileName(string pFileName)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            if (pFileName.IndexOfAny(invalidChars) >= 0)
            {
                pFileName = invalidChars.Aggregate(pFileName, (current, invalidChar) => current.Replace(invalidChar, Convert.ToChar("_")));
            }
            return pFileName;
        }
    }
}
