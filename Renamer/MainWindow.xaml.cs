using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;


namespace Renamer
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            LoadSettings();
        }

        public static bool CheckWildcards;
        public string SelectedAction { get; private set; }
        public string selectedPath { get; private set; } = null;
        public List<(string searchName, string replaceName)> renameList { get; private set; } = new List<(string searchName, string replaceName)>();

        private void LoadSettings()
        {
            this.check_wildcards.IsChecked = Properties.Settings.Default.wildcard;
            for (int i = 1; i <= 10; i++)
            {
                string settingSearch = $"search{i}";
                string settingReplace = $"replace{i}";

                TextBox textBoxSearch = (TextBox)this.FindName(settingSearch);
                TextBox textBoxReplace = (TextBox)this.FindName(settingReplace);

                if (textBoxSearch != null)
                {
                    textBoxSearch.Text = Properties.Settings.Default[settingSearch].ToString();
                }

                if (textBoxReplace != null)
                {
                    textBoxReplace.Text = Properties.Settings.Default[settingReplace].ToString();
                }
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.wildcard = (bool)this.check_wildcards.IsChecked;
            for (int i = 1; i <= 10; i++)
            {
                string settingSearch = $"search{i}";
                string settingReplace = $"replace{i}";

                TextBox textBoxSearch = (TextBox)this.FindName(settingSearch);
                TextBox textBoxReplace = (TextBox)this.FindName(settingReplace);

                if (textBoxSearch != null)
                {
                    Properties.Settings.Default[settingSearch] = textBoxSearch.Text;
                }

                if (textBoxReplace != null)
                {
                    Properties.Settings.Default[settingReplace] = textBoxReplace.Text;
                }
            }
            Properties.Settings.Default.Save();
        }

        public List<(string searchName, string replaceName)> lister()
        {
            List<(string searchName, string replaceName)> renameList = new List<(string searchName, string replaceName)>();

            for (int i = 1; i <= 10; i++)
            {
                string settingSearch = $"search{i}";
                string settingReplace = $"replace{i}";

                TextBox textBoxSearch = (TextBox)this.FindName(settingSearch);
                TextBox textBoxReplace = (TextBox)this.FindName(settingReplace);


                if (textBoxSearch != null && textBoxSearch.Text != "")
                {
                    renameList.Add((textBoxSearch.Text, textBoxReplace.Text));
                }
            }

            return renameList;
        }

        private void startButtonFields_Click(object sender, RoutedEventArgs e)
        {
            CheckWildcards = (bool)this.check_wildcards.IsChecked;
            
            Loader loader = new Loader();
            try
            {
                selectedPath = loader.DirectoryLoader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            if (selectedPath == null)
            {
                MessageBox.Show("Ошибка выбора пути, попробуйте перенести папку на рабочий стол и выбрать её там");
                return;
            }
            
            renameList = lister();

            SelectedAction = "Action1";
            DialogResult = true;
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            CheckWildcards = (bool)this.check_wildcards.IsChecked;

            Loader loader = new Loader();
            try
            {
                selectedPath = loader.DirectoryLoader();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

            if (selectedPath == null)
            {
                MessageBox.Show("Ошибка выбора пути, попробуйте перенести папку на рабочий стол и выбрать её там");
                return;
            }

            renameList = lister();

            SelectedAction = "Action2";
            DialogResult = true;
        }
    }
}
