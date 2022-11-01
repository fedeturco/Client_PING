using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Win32;

namespace Client_PING
{
    /// <summary>
    /// Interaction logic for ExportProfile.xaml
    /// </summary>
    public partial class ExportProfile : Window
    {
        MainWindow main;
        IPEntry selected = null;

        public ExportProfile(MainWindow main_)
        {
            main = main_;

            InitializeComponent();
        }

        private void ButtonExport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "Export_" + DateTime.Now.ToString("yyyy_MM_dd-HH_mm_ss") + ".csv";
            saveFileDialog.Filter = "csv files (*.csv)|*.csv|txt files (*.txt)|*.txt|All files (*.*)|*.*";

            if ((bool)saveFileDialog.ShowDialog())
            {
                string template = TextBoxExportTemplate.Text;
                string fileContent = "";

                foreach (IPEntry ip in main.ListDevices)
                {
                    selected = ip;

                    fileContent += ApplyMacros(template);
                    fileContent += "\n";
                }

                File.WriteAllText(saveFileDialog.FileName, fileContent);
            }
        }

        public string ApplyMacros(string input_)
        {
            string output_ = input_;

            if (selected != null)
            {
                if (selected.CustomArg1 != null)
                    output_ = output_.Replace("<CustomArg1>", selected.CustomArg1);

                if (selected.CustomArg2 != null)
                    output_ = output_.Replace("<CustomArg2>", selected.CustomArg2);

                if (selected.CustomArg3 != null)
                    output_ = output_.Replace("<CustomArg3>", selected.CustomArg3);

                if (selected.StatusPing != null)
                    output_ = output_.Replace("<Ping>", selected.StatusPing);

                if (selected.IpAddress != null)
                    output_ = output_.Replace("<IpAddress>", selected.IpAddress);

                if (selected.Device != null)
                    output_ = output_.Replace("<Device>", selected.Device);

                if (selected.Description != null)
                    output_ = output_.Replace("<Description>", selected.Description);

                if (selected.Group != null)
                    output_ = output_.Replace("<Group>", selected.Group);

                if (selected.MacAddress != null)
                    output_ = output_.Replace("<MacAddress>", selected.MacAddress);

                if (selected.Notes != null)
                    output_ = output_.Replace("<Notes>", selected.Notes);

                if (selected.User != null)
                    output_ = output_.Replace("<User>", selected.User);

                if (selected.Pass != null)
                    output_ = output_.Replace("<Pass>", selected.Pass);

                if (selected.LastPing != null)
                    output_ = output_.Replace("<LastPing>", selected.LastPing);
            }

            return output_;
        }

        private void TextBoxExportTemplate_TextChanged(object sender, TextChangedEventArgs e)
        {
            main.TextBoxExportProfile.Text = this.TextBoxExportTemplate.Text;
        }

        private void AddKeyword(String keyword)
        {
            int pos = TextBoxExportTemplate.SelectionStart;
            string content = TextBoxExportTemplate.Text;

            if((pos) == content.Length)
                TextBoxExportTemplate.Text = content.Substring(0, pos) + keyword;
            else
                TextBoxExportTemplate.Text = content.Substring(0, pos) + keyword + content.Substring(pos);

            TextBoxExportTemplate.SelectionStart = pos + keyword.Length;
            TextBoxExportTemplate.Focus();
        }

        private void LabelPing_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<Ping>");
        }

        private void LabelIpAddress_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<IpAddress>");
        }

        private void LabelDevice_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<Device>");
        }

        private void LabelDescription_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<Description>");
        }

        private void LabelGroup_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<Group>");
        }

        private void LabelUser_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<User>");
        }

        private void LabelPass_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<Pass>");
        }

        private void LabelMacAddress_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<MacAddress>");
        }

        private void LabelLastPing_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<Ping>");
        }

        private void LabelNotes_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<Notes>");
        }

        private void LabelCustomArg1_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<CustomArg1>");
        }

        private void LabelCustomArg2_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<CustomArg2>");
        }

        private void LabelCustomArg3_MouseDown(object sender, MouseButtonEventArgs e)
        {
            AddKeyword("<CustomArg3>");
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.TextBoxExportTemplate.Text = main.TextBoxExportProfile.Text;
        }
    }
}
