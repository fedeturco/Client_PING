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

namespace Client_PING
{
    /// <summary>
    /// Interaction logic for NewProfile.xaml
    /// </summary>
    public partial class NewProfile : Window
    {
        public string profileName = "";

        public NewProfile(string title, string oldName, string buttonName)
        {
            InitializeComponent();

            this.Title = title;
            this.TextBoxProfileName.Text = oldName;
            this.ButtonSaveProfile.Content = buttonName;
            this.TextBoxProfileName.ScrollToEnd();
        }

        private void ButtonSaveProfile_Click(object sender, RoutedEventArgs e)
        {
            this.profileName = TextBoxProfileName.Text;
            this.Close();
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Enter)
            {
                TextBoxProfileName.Text = TextBoxProfileName.Text.Replace("\n", "");
                TextBoxProfileName.Text = TextBoxProfileName.Text.Replace("\r", "");

                ButtonSaveProfile_Click(null, null);
            }

            if(e.Key == Key.Escape)
            {
                this.Close();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            TextBoxProfileName.Focus();
        }
    }
}
