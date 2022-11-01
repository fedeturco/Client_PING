
// -------------------------------------------------------------------------------------------

// Copyright (c) 2022 Federico Turco

// Permission is hereby granted, free of charge, to any person
// obtaining a copy of this software and associated documentation
// files (the "Software"), to deal in the Software without
// restriction, including without limitation the rights to use,
// copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the
// Software is furnished to do so, subject to the following
// conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
// OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
// HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
// FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
// OTHER DEALINGS IN THE SOFTWARE.

// -------------------------------------------------------------------------------------------

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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Web.Script.Serialization;  // System.Web.Extensions
using System.Net;
using Microsoft.Win32;
using System.Net.Sockets;
using System.Net.NetworkInformation;
using System.Threading;
using System.Reflection;

// Visualizza/Nascondi console
using System.Runtime.InteropServices;

namespace Client_PING
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<IPEntry> ListDevices = new ObservableCollection<IPEntry>();

        // Elementi per visualizzare/nascondere la finestra della console
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        bool consoleIsOpen = false;

        // Disable Console Exit Button
        [DllImport("user32.dll")]
        static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        static extern IntPtr DeleteMenu(IntPtr hMenu, uint uPosition, uint uFlags);

        const uint SC_CLOSE = 0xF060;
        const uint MF_BYCOMMAND = (uint)0x00000000L;

        public bool pingRunning = false;
        public bool stopPing = true;
        string oldProfile = "";
        public bool disableSelectProfile = false;

        Thread ThreadPing;
        public bool useThreadsForPing = false;
        IPEntry selected = null;

        public int timeoutPing = 0;
        public int timeoutLastPing = 0;
        public int delayBetweenLoops = 500;

        public bool editedDescriptionOrNotes = false;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void WindowMenu_GuidaPDF_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(Directory.GetCurrentDirectory() + "\\Manuals"))
            {
                // Guida PDF
                var myFiles = Directory.GetFiles(Directory.GetCurrentDirectory() + "\\Manuals");

                String path_manuale = "";

                // debug
                for (int i = 0; i < myFiles.Length; i++)
                {
                    if (myFiles[i].IndexOf(".pdf") != -1)
                    {
                        path_manuale = myFiles[i];
                    }
                }

                if (path_manuale.Length > 0)
                {
                    try
                    {
                        // Eventuale processo con parametro per adobe
                        System.Diagnostics.Process.Start(path_manuale + "#zoom=75");
                    }
                    catch
                    {
                        try
                        {
                            // Nel caso in cui il viewer sia qualco'altro provo ad aprire direttamente il percorso
                            System.Diagnostics.Process.Start(path_manuale);
                        }
                        catch
                        {
                            // Se non trovo niente invio un messaggio di errore
                            MessageBox.Show("Manuale non trovato.", "Info");
                        }
                    }
                }
            }
        }

        // Visualizza console programma da menu tendina
        public void apriConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_SHOW);

            consoleIsOpen = true;
        }

        // Nasconde console programma da menu tendina
        public void chiudiConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            consoleIsOpen = false;
        }

        private void MenuItemApriConsole(object sender, RoutedEventArgs e)
        {
            apriConsole();
        }

        private void MenuItemChiudiConsole(object sender, RoutedEventArgs e)
        {
            chiudiConsole();
        }

        private void MenuItemSalva(object sender, RoutedEventArgs e)
        {
            // Save actual configuration
            saveConfiguration();
        }

        private void MenuItemExit(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        public void saveConfiguration()
        {
            // Salvo Width e Height della finestra
            TextBoxWindowWidth.Text = this.Width.ToString();
            TextBoxWindowHeight.Text = this.Height.ToString();
            CheckBoxWindowMaximized.IsChecked = this.WindowState == WindowState.Maximized;

            // Save current session
            saveCurrentSession();

            // Save current profile content
            if (ComboBoxProfileSelected.SelectedValue != null)
                saveProfile(ComboBoxProfileSelected.SelectedValue.ToString());
        }

        private void saveCurrentSession()
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();

            dynamic toSave = jss.DeserializeObject(File.ReadAllText(Directory.GetCurrentDirectory() + "\\Config\\SettingsToSave.json"));

            Dictionary<string, Dictionary<string, object>> file_ = new Dictionary<string, Dictionary<string, object>>();

            foreach (KeyValuePair<string, object> row in toSave["toSave"])
            {
                // row.key = "textBoxes"
                // row.value = {  }

                switch (row.Key)
                {
                    case "textBoxes":

                        Dictionary<string, object> textBoxes = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            // sub.key = "textBoxModbusAddess_1"
                            // sub.value = { "..." }

                            // debug
                            //Console.WriteLine("sub.key: " + sub.Key);
                            //Console.WriteLine("sub.value: " + sub.Value);

                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                // prop.key = "key"
                                // prop.value = "nomeVariabile"

                                // debug
                                //Console.WriteLine("prop.key: " + prop.Key);
                                //Console.WriteLine("prop.value: " + prop.Value as String);

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        textBoxes.Add(prop.Value as String, (this.FindName(sub.Key) as TextBox).Text);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    textBoxes.Add(sub.Key, (this.FindName(sub.Key) as TextBox).Text);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("textBoxes", textBoxes);
                        break;

                    case "checkBoxes":

                        Dictionary<string, object> checkBoxes = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        checkBoxes.Add(prop.Value as String, (bool)(this.FindName(sub.Key) as CheckBox).IsChecked);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    checkBoxes.Add(sub.Key, (bool)(this.FindName(sub.Key) as CheckBox).IsChecked);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("checkBoxes", checkBoxes);
                        break;

                    case "menuItems":

                        Dictionary<string, object> menuItems = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        menuItems.Add(prop.Value as String, (bool)(this.FindName(sub.Key) as MenuItem).IsChecked);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    menuItems.Add(sub.Key, (bool)(this.FindName(sub.Key) as MenuItem).IsChecked);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("menuItems", menuItems);
                        break;

                    case "radioButtons":

                        Dictionary<string, object> radioButtons = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        radioButtons.Add(prop.Value as String, (this.FindName(sub.Key) as RadioButton).IsChecked);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    radioButtons.Add(sub.Key, (this.FindName(sub.Key) as RadioButton).IsChecked);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("radioButtons", radioButtons);
                        break;

                    case "comboBoxes":

                        Dictionary<string, object> comboBoxes = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        comboBoxes.Add(prop.Value as String, (this.FindName(sub.Key) as ComboBox).SelectedIndex);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    comboBoxes.Add(sub.Key, (this.FindName(sub.Key) as ComboBox).SelectedIndex);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        file_.Add("comboBoxes", comboBoxes);
                        break;

                    case "columns":

                        Dictionary<string, object> columns = new Dictionary<string, object>();

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            // sub.key = "textBoxModbusAddess_1"
                            // sub.value = { "..." }

                            // debug
                            //Console.WriteLine("sub.key: " + sub.Key);
                            //Console.WriteLine("sub.value: " + sub.Value);

                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {
                                // prop.key = "key"
                                // prop.value = "nomeVariabile"

                                // debug
                                //Console.WriteLine("prop.key: " + prop.Key);
                                //Console.WriteLine("prop.value: " + prop.Value as String);

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        columns.Add(prop.Value as String, (this.FindName(sub.Key) as DataGridTextColumn).Width.Value);
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    columns.Add(sub.Key, (this.FindName(sub.Key) as DataGridTextColumn).Width.Value);
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }


                        file_.Add("columns", columns);
                        break;
                }
            }

            File.WriteAllText(Directory.GetCurrentDirectory() + "\\Config\\Config.json", jss.Serialize(file_));
        }

        private bool saveProfile(string name)
        {
            try
            {
                // Salvo tabella PLC
                File_IP_Config save = new File_IP_Config();

                // Saving ip address list
                save.ip_addresses = ListDevices.ToArray<IPEntry>();

                JavaScriptSerializer jss = new JavaScriptSerializer();

                string file_content = jss.Serialize(save);

                File.WriteAllText(Directory.GetCurrentDirectory() + "\\Profiles\\" + name + ".json", file_content);

                return true;
            }
            catch (Exception err)
            {
                Console.WriteLine(err);

                return false;
            }
        }

        public void loadConfiguration()
        {
            // Load all profiles
            loadAllProfiles();

            // Load old configuration
            loadPreviousSession();

            if((bool)CheckBoxCloseConsoleAtStartup.IsChecked)
            {
                chiudiConsole();
            }

            // Salvo Width,Height e WindowState della finestra
            if ((bool)CheckBoxWindowMaximized.IsChecked)
            {
                this.WindowState = WindowState.Maximized;
            }
            else
            {
                int tmp_ = 0;

                if (int.TryParse(TextBoxWindowWidth.Text, out tmp_))
                {
                    this.Width = tmp_;
                }
                if (int.TryParse(TextBoxWindowHeight.Text, out tmp_))
                {
                    this.Height = tmp_;
                }
            }
        }

        public void loadAllProfiles()
        {
            ComboBoxProfileSelected.Items.Clear();

            foreach (String fileName in Directory.GetFiles(Directory.GetCurrentDirectory() + "//Profiles"))
            {
                ComboBoxProfileSelected.Items.Add(System.IO.Path.GetFileNameWithoutExtension(fileName));
            }

            if(ComboBoxProfileSelected.Items.Count > 0)
            {
                ComboBoxProfileSelected.SelectedIndex = 0;
            }
        }
       
        public void loadPreviousSession()
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();

            dynamic toSave = null;
            dynamic loaded = null;

            if (!File.Exists(Directory.GetCurrentDirectory() + "\\Config\\SettingsToSave.json"))
            {
                Console.WriteLine("File: {0} not found", Directory.GetCurrentDirectory() + "\\Config\\SettingsToSave.json");
                return;
            }

            toSave = jss.DeserializeObject(File.ReadAllText(Directory.GetCurrentDirectory() + "\\Config\\SettingsToSave.json"));

            if (!File.Exists(Directory.GetCurrentDirectory() + "\\Config\\Config.json"))
            {
                Console.WriteLine("File: {0} not found", Directory.GetCurrentDirectory() + "\\Config\\Config.json");
                return;
            }

            loaded = jss.DeserializeObject(File.ReadAllText(Directory.GetCurrentDirectory() + "\\Config\\Config.json"));

            foreach (KeyValuePair<string, object> row in toSave["toSave"])
            {
                // row.key = "textBoxes"
                // row.value = {  }

                switch (row.Key)
                {
                    case "textBoxes":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        // debug
                                        //Console.WriteLine(" -- ");
                                        //Console.WriteLine("sub.Key: " + sub.Key);
                                        //Console.WriteLine("prop.Value: " + prop.Value);
                                        //Console.WriteLine(loaded[row.Key][prop.Value.ToString()]);

                                        try
                                        {
                                            (this.FindName(sub.Key) as TextBox).Text = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as TextBox).Text = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "checkBoxes":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as CheckBox).IsChecked = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as CheckBox).IsChecked = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "menuItems":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as MenuItem).IsChecked = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as MenuItem).IsChecked = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "radioButtons":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as RadioButton).IsChecked = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as RadioButton).IsChecked = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "comboBoxes":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        try
                                        {
                                            (this.FindName(sub.Key) as ComboBox).SelectedIndex = loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as ComboBox).SelectedIndex = loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;

                    case "columns":

                        foreach (KeyValuePair<string, object> sub in toSave["toSave"][row.Key])
                        {
                            bool found = false;

                            foreach (KeyValuePair<string, object> prop in toSave["toSave"][row.Key][sub.Key])
                            {

                                if (prop.Key == "key")
                                {
                                    found = true;

                                    if (this.FindName(sub.Key) != null)
                                    {
                                        // debug
                                        //Console.WriteLine(" -- ");
                                        //Console.WriteLine("sub.Key: " + sub.Key);
                                        //Console.WriteLine("prop.Value: " + prop.Value);
                                        //Console.WriteLine(loaded[row.Key][prop.Value.ToString()]);

                                        try
                                        {
                                            (this.FindName(sub.Key) as DataGridTextColumn).Width = (double)loaded[row.Key][prop.Value.ToString()];
                                        }
                                        catch (Exception err)
                                        {
                                            Console.WriteLine(prop.Value.ToString() + " generated an error");
                                            Console.WriteLine(err);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine(sub.Key + " not found in current form");
                                    }
                                }
                            }

                            if (!found)
                            {
                                if (this.FindName(sub.Key) != null)
                                {
                                    try
                                    {
                                        (this.FindName(sub.Key) as DataGridTextColumn).Width = (double)loaded[row.Key][sub.Key.ToString()];
                                    }
                                    catch (Exception err)
                                    {
                                        Console.WriteLine(sub.Key.ToString() + " generated an error");
                                        Console.WriteLine(err);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine(sub.Key + " not found in current form");
                                }
                            }
                        }

                        break;
                }
            }
        }

        public void loadProfile(string name)
        {
            string profilePath = Directory.GetCurrentDirectory() + "\\Profiles\\" + System.IO.Path.GetFileNameWithoutExtension(name) + ".json";
            
            if (File.Exists(profilePath))
            {
                ListDevices.Clear();

                JavaScriptSerializer jss = new JavaScriptSerializer();

                File_IP_Config saved = new File_IP_Config();

                saved = jss.Deserialize<File_IP_Config>(File.ReadAllText(Directory.GetCurrentDirectory() + "\\Profiles\\" + System.IO.Path.GetFileNameWithoutExtension(name) + ".json"));

                for (int i = 0; i < saved.ip_addresses.Length; i++)
                {
                    ListDevices.Add(saved.ip_addresses[i]);
                }

                LabelDeviceCount.Dispatcher.Invoke((Action)delegate
                {
                    LabelDeviceCount.Content = saved.ip_addresses.Length.ToString();
                });

                Console.WriteLine("Loaded profile: " + profilePath);
            }
            else
            {
                Console.WriteLine("Selected profile doesn't exist: " + profilePath);
            }
        }

        private void ButtonPing_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (selected != null)
                {
                    Process p = new Process();

                    p.StartInfo = new ProcessStartInfo()
                    {
                        WorkingDirectory = Directory.GetCurrentDirectory(),
                        FileName = "ping",
                        Arguments = "-t " + selected.IpAddress + " -w " + timeoutPing.ToString(),
                        UseShellExecute = true,
                        RedirectStandardOutput = false,
                        RedirectStandardError = false,
                        CreateNoWindow = false
                    };

                    p.Start();
                }
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb0_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathBrowser.Text,
                    Arguments = ApplyMacros(TextBoxPathBrowser_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        public string ApplyMacros(string input_)
        {
            string output_ = input_;

            if (selected != null) 
            {
                if(selected.CustomArg1 != null)
                    output_ = output_.Replace("<CustomArg1>", selected.CustomArg1);

                if (selected.CustomArg2 != null)
                    output_ = output_.Replace("<CustomArg2>", selected.CustomArg2);

                if (selected.CustomArg3 != null)
                    output_ = output_.Replace("<CustomArg3>", selected.CustomArg3);

                if (selected.Device != null)
                    output_ = output_.Replace("<Device>", selected.Device);

                if (selected.IpAddress != null)
                    output_ = output_.Replace("<IpAddress>", selected.IpAddress);

                if (selected.User != null)
                    output_ = output_.Replace("<User>", selected.User);

                if (selected.Pass != null)
                    output_ = output_.Replace("<Pass>", selected.Pass);

                if (selected.MacAddress != null)
                    output_ = output_.Replace("<MacAddress>", selected.MacAddress);
            }

            return output_;
        }

        private void ButtonCopyWeb0_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(TextBoxIpAddress.Text.ToString());
        }

        private void ButtonCopyWeb1_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(TextBoxCustomLink1.Text.ToString());
        }

        private void ButtonCopyWeb2_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(TextBoxCustomLink2.Text.ToString());
        }

        private void ButtonCopyWeb3_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(TextBoxCustomLink3.Text.ToString());
        }

        private void ButtonLoop_Click(object sender, RoutedEventArgs e)
        {
            if (stopPing)
            {
                stopPing = false;
                ButtonLoopTextBlock.Text = "Stop";
                
                ButtonLoopSingle.IsEnabled = false;
                CheckBoxUseThreadsForPing.IsEnabled = false;
                ComboBoxProfileSelected.IsEnabled = false;
                ButtonAddNewProfile.IsEnabled = false;

                ThreadPing = new Thread(new ThreadStart(AskPing));
                ThreadPing.IsBackground = true;
                ThreadPing.Start();
            }
            else
            {
                stopPing = true;

                ButtonLoop.IsEnabled = false;
                ButtonLoopSingle.IsEnabled = false;
                // CheckBoxUseThreadsForPing.IsEnabled = false;
                // ComboBoxProfileSelected.IsEnabled = true;
            }
        }

        private void ButtonLoopSingle_Click(object sender, RoutedEventArgs e)
        {
            stopPing = true;

            ButtonLoop.IsEnabled = false;
            ButtonLoopSingle.IsEnabled = false;
            CheckBoxUseThreadsForPing.IsEnabled = false;
            ComboBoxProfileSelected.IsEnabled = false;
            ButtonAddNewProfile.IsEnabled = false;

            Thread t = new Thread(new ThreadStart(AskPing));
            t.Start();
        }

        public void AskPing()
        {
            pingRunning = true;

            do
            {
                this.Dispatcher.Invoke((Action)delegate
                {
                    BorderLoopRunning.Background = Brushes.Yellow;
                });

                PingAll();

                if (!stopPing)
                {
                    Thread.Sleep(delayBetweenLoops);
                }

            }
            while (!stopPing);

            ButtonLoop.Dispatcher.Invoke((Action)delegate
            {
                ButtonLoopTextBlock.Text = "Start";

                ButtonLoop.IsEnabled = true;
                ButtonLoopSingle.IsEnabled = true;
                CheckBoxUseThreadsForPing.IsEnabled = true;
                ComboBoxProfileSelected.IsEnabled = true;
                ButtonAddNewProfile.IsEnabled = true;

                DataGridListaIp.IsEnabled = true;
                
                BorderLoopRunning.Background = SystemColors.ControlBrush;
            });

            pingRunning = false;
        }

        public void PingAll()
        {
            try
            {
                foreach (IPEntry ipTmp in ListDevices)
                {
                    ipTmp.ColorPing = Brushes.White.ToString();
                }

                this.Dispatcher.Invoke((Action)delegate
                {
                    DataGridListaIp.ItemsSource = null;
                    DataGridListaIp.ItemsSource = ListDevices;
                });
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }

            if (useThreadsForPing)
                Console.WriteLine("Starting ping thread");

            try
            {
                foreach (IPEntry ipTmp in ListDevices)
                {
                    Console.WriteLine(ipTmp.IpAddress);

                    if (!useThreadsForPing)
                    {
                        PingIp(ipTmp);
                        Thread.Sleep(100);
                    }
                    else
                    {
                        Thread t = new Thread(new ParameterizedThreadStart(PingIp));
                        t.Start(ipTmp);
                    }
                }

                if (useThreadsForPing)
                {
                    // Thread.Sleep(3 * timeoutPing + 500);

                    int wait = 3 * timeoutPing + 500;
                    
                    for(int i = 0; i < (int)wait/100; i++)
                    {
                        Thread.Sleep(100);

                        if(i * 100 * 2 > timeoutPing)
                            if (stopPing)
                                return;
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }
        }

        public void PingIp(object obj)
        {
            IPEntry ipTmp = (IPEntry)obj;
            Ping pingSender = new Ping();
            PingReply RispostaPING;

            if (ipTmp.IpAddress != null)
            {
                if (ipTmp.IpAddress.Length > 0)
                {
                    RispostaPING = pingSender.Send(ipTmp.IpAddress, timeoutPing);

                    if (RispostaPING.Status != IPStatus.Success)
                    {
                        RispostaPING = pingSender.Send(ipTmp.IpAddress, timeoutPing);

                        if (RispostaPING.Status != IPStatus.Success)
                        {
                            ipTmp.StatusPing = "Timeout";
                            ipTmp.ColorPing = Brushes.Red.ToString();
                        }
                        else
                        {
                            ipTmp.StatusPing = "OK - " + RispostaPING.RoundtripTime + " ms";
                            ipTmp.LastPing = DateTime.Now.ToString();
                            ipTmp.ColorPing = Brushes.Lime.ToString();
                        }
                    }
                    else
                    {
                        ipTmp.StatusPing = "OK - " + RispostaPING.RoundtripTime + " ms";
                        ipTmp.LastPing = DateTime.Now.ToString();
                        ipTmp.ColorPing = Brushes.Lime.ToString();
                    }

                    try
                    {
                        TimeSpan diff = DateTime.Now - Convert.ToDateTime(ipTmp.LastPing);

                        if (diff.TotalSeconds > timeoutLastPing)
                        {
                            ipTmp.ColorLastPing = Brushes.Red.ToString();
                        }
                        else
                        {
                            ipTmp.ColorLastPing = Brushes.Lime.ToString();
                        }
                    }
                    catch(Exception err)
                    {
                        // debug
                        Console.WriteLine(err);

                        ipTmp.ColorLastPing = Brushes.Red.ToString();
                    }

                    this.Dispatcher.Invoke((Action)delegate
                    {
                        DataGridListaIp.ItemsSource = null;
                        DataGridListaIp.ItemsSource = ListDevices;
                    });
                }
            }
        }

        private void WindowMenu_Info_Click(object sender, RoutedEventArgs e)
        {
            Info window = new Info(this.Title.Split('-')[0], GetBuildVersion());
            window.Show();
        }

        private void DataGridListaIp_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateIpSelected();    
        }

        public void UpdateIpSelected()
        {
            try
            {
                /*LabelIpDevice.Content = "???";

                TextBoxIpAddress.Text = "???";
                TextBoxMacAddress.Text = "???";

                TextBoxCustomLink1.Text = "???";
                TextBoxCustomLink2.Text = "???";
                TextBoxCustomLink3.Text = "???";

                RichTextBoxNotes.Document.Blocks.Clear();*/

                MenuBar.Background = Brushes.White;

                if ((IPEntry)(DataGridListaIp.SelectedItem) != null)
                    selected = (IPEntry)(DataGridListaIp.SelectedItem);

                if (selected != null)
                {
                    if(selected.Device != null)
                        LabelIpDevice.Content = selected.Device;
                    else
                        LabelIpDevice.Content = "";

                    if(selected.IpAddress != null)
                        TextBoxIpAddress.Text = selected.IpAddress;
                    else
                        TextBoxIpAddress.Text = "";

                    if (selected.MacAddress != null)
                        TextBoxMacAddress.Text = selected.MacAddress;
                    else
                        TextBoxMacAddress.Text = "";

                    if (selected.CustomArg1 != null)
                        TextBoxCustomLink1.Text = selected.CustomArg1;

                    if(selected.CustomArg2 != null)
                        TextBoxCustomLink2.Text = selected.CustomArg2;

                    if(selected.CustomArg3 != null)
                        TextBoxCustomLink3.Text = selected.CustomArg3;

                    if (selected.Device != null && selected.IpAddress != null)
                        LabelCurrentIp.Content = selected.Device + ": " + selected.IpAddress;

                    else if(selected.IpAddress != null)
                        LabelCurrentIp.Content = selected.IpAddress;

                    if (selected.Notes != null)
                    {
                        RichTextBoxNotes.Document.Blocks.Clear();
                        //RichTextBoxNotes.AppendText("\n");
                        RichTextBoxNotes.AppendText(selected.Notes);
                    }
                    else
                    {
                        RichTextBoxNotes.Document.Blocks.Clear();
                    }

                    if (selected.Description != null)
                    {
                        RichTextBoxDescription.Document.Blocks.Clear();
                        //RichTextBoxDescription.AppendText("\n");
                        RichTextBoxDescription.AppendText(selected.Description);
                    }
                    else
                    {
                        RichTextBoxDescription.Document.Blocks.Clear();
                    }

                    enableButtons(true);
                }
            }
            catch (Exception err)
            {
                // Debug
                Console.WriteLine(err);

                enableButtons(false);
            }
        }

        public void enableButtons(bool enable)
        {
            try
            {
                ButtonPing_Main.IsEnabled = enable;

                ButtonWeb0.IsEnabled = enable;

                ButtonWeb0_Main.IsEnabled = enable;
                
                ButtonWeb1_Main.IsEnabled = enable;
                ButtonWeb2_Main.IsEnabled = enable;
                ButtonWeb3_Main.IsEnabled = enable;
                ButtonWeb4_Main.IsEnabled = enable;
                ButtonWeb5_Main.IsEnabled = enable;
                ButtonWeb6_Main.IsEnabled = enable;
                ButtonWeb7_Main.IsEnabled = enable;
                ButtonWeb8_Main.IsEnabled = enable;
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }
        }

        public string GetBuildVersion()
        {
            // Info build version
            var version = Assembly.GetExecutingAssembly().GetName().Version;

            var buildDate = new DateTime(2000, 1, 1)
               .AddDays(version.Build).AddSeconds(version.Revision * 2);

            var displayableVersion = $"{version} ({buildDate})";

            string output = displayableVersion;

            Console.WriteLine("Current version (inc. build date) = " + output);

            return output;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Setto il titolo della finestra
            string build = GetBuildVersion();
            this.Title = "Client PING - v" + build.Split(' ')[0].Split('.')[0] + "." + build.Split(' ')[0].Split('.')[1];

            // Aspetti grafici
            // ColumnStatusPing.Visibility = Visibility.Collapsed;
            // ColumnIpAddress.Visibility = Visibility.Collapsed;
            // ColumnDevice.Visibility = Visibility.Collapsed;
            // ColumnDescription.Visibility = Visibility.Collapsed;
            // ColumnGroup.Visibility = Visibility.Collapsed;
            // ColumnUser.Visibility = Visibility.Collapsed;
            // ColumnPass.Visibility = Visibility.Collapsed;
            // ColumnMacAddress.Visibility = Visibility.Collapsed;
            // ColumnNotes.Visibility = Visibility.Collapsed;
            // ColumnCustomArg1.Visibility = Visibility.Collapsed;
            // ColumnCustomArg2.Visibility = Visibility.Collapsed;
            // ColumnCustomArg3.Visibility = Visibility.Collapsed;
            // ColumnLastPing.Visibility = Visibility.Collapsed;

            ButtonPing_Main.IsEnabled = false;
            ButtonWeb0_Main.IsEnabled = false;
            ButtonWeb1_Main.IsEnabled = false;
            ButtonWeb2_Main.IsEnabled = false;
            ButtonWeb3_Main.IsEnabled = false;
            ButtonWeb4_Main.IsEnabled = false;
            ButtonWeb5_Main.IsEnabled = false;
            ButtonWeb6_Main.IsEnabled = false;
            ButtonWeb7_Main.IsEnabled = false;
            ButtonWeb8_Main.IsEnabled = false;

            // Carico la configurazioe precedente
            loadConfiguration();

            // Assegno il datasource alla tabella
            DataGridListaIp.ItemsSource = null;
            DataGridListaIp.ItemsSource = ListDevices;
        }

        private void ComboBoxProfileSelected_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComboBoxProfileSelected.SelectedItem != null && !disableSelectProfile)
            {
                // debug
                // Console.WriteLine("selected: " + ComboBoxProfileSelected.SelectedValue);

                // Salvo il profilo precedente
                if (oldProfile.Length > 0)
                    saveProfile(oldProfile);

                // Carico il profilo selezionato
                loadProfile(ComboBoxProfileSelected.SelectedValue.ToString());

                // Aggiorno la variabile oldProfile
                oldProfile = ComboBoxProfileSelected.SelectedValue.ToString();
            }
        }

        private void ButtonAddNewProfile_Click(object sender, RoutedEventArgs e)
        {
            NewProfile window = new NewProfile("New profile", "", "Save");
            window.ShowDialog();

            ComboBoxProfileSelected.Items.Add(window.profileName);
            ComboBoxProfileSelected.SelectedValue = window.profileName;

            ListDevices.Clear();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Fermo il ping in ogni caso
            stopPing = true;

            // Save actual configuration
            saveConfiguration();
        }

        private void MenuItemImportCsv(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Filter = "csv files (*.csv)|*.csv|txt files (*.txt)|*.txt|All files (*.*)|*.*";

            if ((bool)openFileDialog.ShowDialog())
            {
                try
                {
                    string content_ = File.ReadAllText(openFileDialog.FileName);
                    content_ = content_.Replace("\r", "");

                    ListDevices.Clear();

                    foreach (string line_ in content_.Split('\n'))
                    {
                        IPEntry IPEntry_ = new IPEntry();

                        if (line_.Split(',').Length == 0)
                            continue;

                        IPEntry_.IpAddress = line_.Split(',')[0];

                        if (line_.Split(',').Length > 1)
                            IPEntry_.Device = line_.Split(',')[1];

                        if (line_.Split(',').Length > 2)
                            IPEntry_.Description = line_.Split(',')[2];

                        ListDevices.Add(IPEntry_);
                    }
                }
                catch (Exception err)
                {
                    // debug
                    Console.WriteLine(err);

                    MessageBox.Show("Error importing csv file", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);

                    // In caso di errori cancello la lista
                    ListDevices.Clear();
                }
            }
        }

        private void changeColumnVisibility(object sender, RoutedEventArgs e)
        {
            DataGridTextColumn toEdit = (DataGridTextColumn)this.FindName((sender as MenuItem).Name.Replace("View", "Column"));

            if (toEdit != null)
            {
                if ((sender as MenuItem).IsChecked)
                {
                    toEdit.Visibility = Visibility.Visible;
                }
                else
                {
                    toEdit.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            // Tasto Ctrl premuto
            if(Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                if(e.Key == Key.P)
                {
                    ButtonPing_Main_Click(null, null);
                }

                if (e.Key == Key.K)
                {
                    ButtonLoopSingle_Click(null, null);
                }
                if (e.Key == Key.L)
                {
                    ButtonLoop_Click(null, null);
                }

                if (e.Key == Key.W)
                {
                    ButtonWeb0_Main_Click(null, null);
                }
                if (e.Key == Key.D1)
                {
                    ButtonWeb1_Main_Click(null, null);
                }
                if (e.Key == Key.D2)
                {
                    ButtonWeb2_Main_Click(null, null);
                }
                if (e.Key == Key.D3)
                {
                    ButtonWeb3_Main_Click(null, null);
                }
                if (e.Key == Key.D4)
                {
                    ButtonWeb4_Main_Click(null, null);
                }
                if (e.Key == Key.D5)
                {
                    ButtonWeb5_Main_Click(null, null);
                }
                if (e.Key == Key.D6)
                {
                    ButtonWeb6_Main_Click(null, null);
                }
                if (e.Key == Key.D7)
                {
                    ButtonWeb7_Main_Click(null, null);
                }
                if (e.Key == Key.D8)
                {
                    ButtonWeb8_Main_Click(null, null);
                }

                if (e.Key == Key.I)
                {
                    WindowMenu_Info_Click(null, null);
                }

                // Cambia tab
                if (e.Key == Key.D1)
                {
                    TabControlGenerale.SelectedIndex = 0;
                }
                if (e.Key == Key.D2)
                {
                    TabControlGenerale.SelectedIndex = 1;
                }
                if (e.Key == Key.D3)
                {
                    TabControlGenerale.SelectedIndex = 2;
                }

                // Tasto Shift premuto
                if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)) 
                {
                    // Apri/chiudi console
                    if (e.Key == Key.C)
                    {
                        if (!consoleIsOpen)
                        {
                            apriConsole();
                            this.Focus();
                        }
                        else
                        {
                            chiudiConsole();
                        }
                    }
                }
            }
        }

        private void TextBoxLabelApp1_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(TextBoxLabelApp1.Text.Length > 0)
            {
                ButtonWeb1_Main_Text.Text = TextBoxLabelApp1.Text;
                ButtonWeb1_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb1_Main.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxLabelApp2_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(TextBoxLabelApp2.Text.Length > 0)
            {
                ButtonWeb2_Main_Text.Text = TextBoxLabelApp2.Text;
                ButtonWeb2_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb2_Main.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxLabelApp3_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(TextBoxLabelApp3.Text.Length > 0)
            {
                ButtonWeb3_Main_Text.Text = TextBoxLabelApp3.Text;
                ButtonWeb3_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb3_Main.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxTimeout_TextChanged(object sender, TextChangedEventArgs e)
        {
            int out_ = 0;

            if (int.TryParse(TextBoxTimeout.Text, out out_)) 
            {
                timeoutPing = out_;
            }
            else
            {
                TextBoxTimeout.Text = "200";
            }
        }

        private void TextBoxTimeoutLastPing_TextChanged(object sender, TextChangedEventArgs e)
        {
            int out_ = 0;

            if (int.TryParse(TextBoxTimeoutLastPing.Text, out out_))
            {
                timeoutLastPing = out_;
            }
            else
            {
                TextBoxTimeoutLastPing.Text = "60";
            }
        }

        private void CheckBoxPinWIndow_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = (bool)CheckBoxPinWIndow.IsChecked;
        }

        private void ButtonWeb1_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp1.Text,
                    Arguments = ApplyMacros(TextBoxPathApp1_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb2_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp2.Text,
                    Arguments = ApplyMacros(TextBoxPathApp2_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb3_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp3.Text,
                    Arguments = ApplyMacros(TextBoxPathApp3_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb4_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp4.Text,
                    Arguments = ApplyMacros(TextBoxPathApp4_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb5_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp5.Text,
                    Arguments = ApplyMacros(TextBoxPathApp5_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb6_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp6.Text,
                    Arguments = ApplyMacros(TextBoxPathApp6_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb7_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp7.Text,
                    Arguments = ApplyMacros(TextBoxPathApp7_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void ButtonWeb8_Main_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process p = new Process();

                p.StartInfo = new ProcessStartInfo()
                {
                    WorkingDirectory = Directory.GetCurrentDirectory(),
                    FileName = TextBoxPathApp8.Text,
                    Arguments = ApplyMacros(TextBoxPathApp8_Arg0.Text),
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = false
                };

                p.Start();
            }
            catch (Exception error)
            {
                Console.WriteLine(error);
            }
        }

        private void TextBoxLabelApp4_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(TextBoxLabelApp4.Text.Length > 0)
            {
                ButtonWeb4_Main_Text.Text = TextBoxLabelApp4.Text;
                ButtonWeb4_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb4_Main.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxLabelApp5_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(TextBoxLabelApp5.Text.Length > 0)
            {
                ButtonWeb5_Main_Text.Text = TextBoxLabelApp5.Text;
                ButtonWeb5_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb5_Main.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxLabelApp6_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(TextBoxLabelApp6.Text.Length > 0)
            {
                ButtonWeb6_Main_Text.Text = TextBoxLabelApp6.Text;
                ButtonWeb6_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb6_Main.Visibility = Visibility.Hidden;
            }
        }

        private void TextBoxLabelApp7_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TextBoxLabelApp7.Text.Length > 0)
            {
                ButtonWeb7_Main_Text.Text = TextBoxLabelApp7.Text;
                ButtonWeb7_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb7_Main.Visibility = Visibility.Hidden;
            }
        }
        private void TextBoxLabelApp8_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TextBoxLabelApp8.Text.Length > 0)
            {
                ButtonWeb8_Main_Text.Text = TextBoxLabelApp8.Text;
                ButtonWeb8_Main.Visibility = Visibility.Visible;
            }
            else
            {
                ButtonWeb8_Main.Visibility = Visibility.Hidden;
            }
        }

        private void CheckBoxUseThreadsForPing_Checked(object sender, RoutedEventArgs e)
        {
            useThreadsForPing = (bool)CheckBoxUseThreadsForPing.IsChecked;
        }

        private void TextBoxDelayBetweenLoops_TextChanged(object sender, TextChangedEventArgs e)
        {
            int out_ = 0;

            if(int.TryParse(TextBoxDelayBetweenLoops.Text, out out_))
            {
                delayBetweenLoops = out_;
            }
            else
            {
                TextBoxDelayBetweenLoops.Text = "500";
            }
        }

        public void ApplyIconFromExe(Image image, String path)
        {
            if (System.IO.File.Exists(path))
            {
                var sourceIcon = System.Drawing.Icon.ExtractAssociatedIcon(path);
                var bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(sourceIcon.Handle, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions()); sourceIcon.Dispose();

                image.Source = new TransformedBitmap(bitmapSource,
                    new ScaleTransform(
                        image.Width / bitmapSource.Width,
                        image.Height / bitmapSource.Height));
            }
        }

        private void TextBoxPathApp1_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb1_Main_Image, TextBoxPathApp1.Text);
        }

        private void TextBoxPathApp2_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb2_Main_Image, TextBoxPathApp2.Text);
        }

        private void TextBoxPathApp3_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb3_Main_Image, TextBoxPathApp3.Text);
        }

        private void TextBoxPathApp4_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb4_Main_Image, TextBoxPathApp4.Text);
        }

        private void TextBoxPathApp5_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb5_Main_Image, TextBoxPathApp5.Text);
        }

        private void TextBoxPathApp6_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb6_Main_Image, TextBoxPathApp6.Text);
        }

        private void TextBoxPathApp7_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb7_Main_Image, TextBoxPathApp7.Text);
        }

        private void TextBoxPathApp8_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyIconFromExe(ButtonWeb8_Main_Image, TextBoxPathApp8.Text);
        }

        private void ButtonDeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (ComboBoxProfileSelected.SelectedItem != null)
            {
                if (MessageBox.Show("Delete profile " + ComboBoxProfileSelected.SelectedValue.ToString(), "Info", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    disableSelectProfile = true;

                    // Cancello il file json
                    System.IO.File.Delete(Directory.GetCurrentDirectory() + "\\Profiles\\" + System.IO.Path.GetFileNameWithoutExtension(ComboBoxProfileSelected.SelectedValue.ToString()) + ".json");

                    // Load all profiles
                    loadAllProfiles();

                    // Seleziono il primo profilo disponibile
                    if (ComboBoxProfileSelected.Items.Count > 0)
                    {
                        ComboBoxProfileSelected.SelectedIndex = 0;

                        // Carico il profilo selezionato
                        loadProfile(ComboBoxProfileSelected.SelectedValue.ToString());

                        oldProfile = ComboBoxProfileSelected.SelectedValue.ToString();
                    }

                    disableSelectProfile = false;
                }
            }
        }

        private void ButtonRenameProfile_Click(object sender, RoutedEventArgs e)
        {
            if (ComboBoxProfileSelected.SelectedItem != null)
            {
                NewProfile window = new NewProfile("Rename profile", ComboBoxProfileSelected.SelectedValue.ToString(), "Rename");
                window.ShowDialog();

                disableSelectProfile = true;

                File.Move(
                    Directory.GetCurrentDirectory() + "\\Profiles\\" + System.IO.Path.GetFileNameWithoutExtension(ComboBoxProfileSelected.SelectedValue.ToString()) + ".json",
                    Directory.GetCurrentDirectory() + "\\Profiles\\" + System.IO.Path.GetFileNameWithoutExtension(window.profileName) + ".json"
                    );

                // Reload all profile
                loadAllProfiles();

                ComboBoxProfileSelected.SelectedValue = window.profileName;

                disableSelectProfile = false;
            }
        }

        private void ButtonOpenProfileFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start("explorer.exe", Directory.GetCurrentDirectory() + "\\Profiles\\");
            }
            catch
            {
            }
        }

        private void ButtonExportProfile_Click(object sender, RoutedEventArgs e)
        {
            ExportProfile exportProfile = new ExportProfile(this);
            exportProfile.Show();
        }

        private void RichTextBoxDescription_TextChanged(object sender, TextChangedEventArgs e)
        {
            if(TabControlGenerale.SelectedIndex == 1)
            {
                if (selected != null)
                {
                    TextRange tr = new TextRange(RichTextBoxDescription.Document.ContentStart, RichTextBoxDescription.Document.ContentEnd);
                    selected.Description = tr.Text.Substring(0, tr.Text.Length - 2); ;

                    editedDescriptionOrNotes = true;
                }
            }
        }

        private void RichTextBoxNotes_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TabControlGenerale.SelectedIndex == 1)
            {
                if (selected != null)
                {
                    TextRange tr = new TextRange(RichTextBoxNotes.Document.ContentStart, RichTextBoxNotes.Document.ContentEnd);
                    selected.Notes = tr.Text.Substring(0, tr.Text.Length - 2);

                    editedDescriptionOrNotes = true;
                }
            }
        }

        private void ViewResetColumnWidth_Click(object sender, RoutedEventArgs e)
        {
            ColumnStatusPing.Width = 100;
            ColumnIpAddress.Width = 100;
            ColumnDevice.Width = 100;
            ColumnDescription.Width = 100;
            ColumnGroup.Width = 100;
            ColumnUser.Width = 100;
            ColumnPass.Width = 100;
            ColumnMacAddress.Width = 100;
            ColumnNotes.Width = 100;
            ColumnCustomArg1.Width = 100;
            ColumnCustomArg2.Width = 100;
            ColumnCustomArg3.Width = 100;
            ColumnLastPing.Width = 160;
        }

        private void TabControlGenerale_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (TabControlGenerale.SelectedIndex)
            {
                case 0:
                    if (editedDescriptionOrNotes)
                    {
                        editedDescriptionOrNotes = false;

                        DataGridListaIp.ItemsSource = null;
                        DataGridListaIp.ItemsSource = ListDevices;
                    }
                    break;
            }
        }
    }
    public class File_IP_Config
    {
        public IPEntry[] ip_addresses { get; set; }
    }


    public class IPEntry
    {
        public string IpAddress { get; set; }       // IP Address
        public string Device { get; set; }          // Device Info
        public string Description { get; set; }     // Device Description
        public string Group { get; set; }           // Group to sort devices
        public string User { get; set; }            // Username
        public string Pass { get; set; }            // Password
        public string StatusPing { get; set; }      // Status PING
        public string LastPing { get; set; }        // DateTime Last ping
        public string MacAddress { get; set; }      // MACAddress
        public string Notes { get; set; }           // Notes about the device
        public string CustomArg1 { get; set; }     // CustomArg to access something on the device
        public string CustomArg2 { get; set; }     // CustomArg to access something on the device
        public string CustomArg3 { get; set; }     // CustomArg to access something on the device

        // Hidden properties
        public string ColorPing { get; set; }
        public string ColorLastPing { get; set; }

    }
}
