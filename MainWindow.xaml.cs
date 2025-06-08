using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Drawing;
using System.IO;
using MessageBox = System.Windows.MessageBox;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace AutoDocFlow
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        App _app;

        public MainWindow(App app)
        {
            InitializeComponent();
            _app = app;

            DP_startDate.SelectedDate = DateTime.Now.AddMonths(-1);
            DP_endDate.SelectedDate = DateTime.Now;            

            CoB_contractor.Items.Add("Все контрагенты");
            CoB_contractor.SelectedIndex = 0;
        }

        public void UpdatePaths(string dbPath, string docTempPath, string signaturePath, string stampPath, string mailTempPath, string outputPath, string orgName, string orgPersonName, string orgEmail, string orgEmailPassword)
        {
            Dispatcher.Invoke(() =>
            {
                TB_dbPath.Text = dbPath;
                TB_docTemplate.Text = docTempPath;
                TB_signaturePath.Text = signaturePath;
                TB_stampPath.Text = stampPath;
                TB_mailTemplate.Text = mailTempPath;
                TB_outputPath.Text = outputPath;
                TB_orgName.Text = orgName;
                TB_orgPersonName.Text = orgPersonName;
                TB_orgEmail.Text = orgEmail;
                TB_orgEmailPassword.Password = orgEmailPassword;
            });
        }

        public void UpdateContractorsList(List<string> contractorsList)
        {
            Dispatcher.Invoke(() =>
            {
                CoB_contractor.Items.Clear();
                CoB_contractor.ItemsSource = contractorsList;
                CoB_contractor.SelectedIndex = 0;
            });
        }

        public void UpdateDates(DateTime start, DateTime end)
        {
            Dispatcher.Invoke(() =>
            {
                DP_startDate.SelectedDate = start;
                DP_endDate.SelectedDate = end;
            });
        }

        public List<string> GetSettings()
        {
            List<string> paths = new List<string>();
            Dispatcher.Invoke(() => {
                paths.Add(TB_dbPath.Text);
                paths.Add(TB_docTemplate.Text);
                paths.Add(TB_signaturePath.Text);
                paths.Add(TB_stampPath.Text);
                paths.Add(TB_mailTemplate.Text);
                paths.Add(TB_outputPath.Text);
                paths.Add(TB_orgName.Text);
                paths.Add(TB_orgPersonName.Text);
                paths.Add(TB_orgEmail.Text);
                paths.Add(TB_orgEmailPassword.Password);
            });
            return paths;
        }

        public string GetScriptParams(ScriptQueryType queryType)
        {
            string contractor;
            var sb = new StringBuilder();
            sb.Append(queryType.ToString());
            Dispatcher.Invoke(() =>
            {
                switch (queryType)
                {
                    case ScriptQueryType.GENERATE:
                        sb.Append($" \"{TB_dbPath.Text}\" ");
                        sb.Append($"\"{TB_docTemplate.Text}\" ");
                        sb.Append($"\"{TB_signaturePath.Text}\" ");
                        sb.Append($"\"{TB_stampPath.Text}\" ");
                        sb.Append($"\"{TB_mailTemplate.Text}\" ");
                        sb.Append($"\"{TB_outputPath.Text}\" ");
                        sb.Append($"\"{TB_orgName.Text}\" ");
                        sb.Append($"\"{TB_orgPersonName.Text}\" ");
                        sb.Append($"\"{TB_orgEmail.Text}\" ");
                        sb.Append($"\"{TB_orgEmailPassword.Password}\" ");

                        try
                        {
                            sb.Append($"\"{CB_filterPeriod.IsChecked.Value.ToString()}\" ");
                            sb.Append($"\"{DP_startDate.SelectedDate.Value.ToString()[..10]}\" ");
                            sb.Append($"\"{DP_endDate.SelectedDate.Value.ToString()[..10]}\" ");
                        }
                        catch (Exception ex)
                        {
                            throw new Exception($"Неверно введена дата");
                        }

                        sb.Append($"\"{CB_filterContractor.IsChecked.Value.ToString()}\" ");
                        sb.Append($"\"{CoB_contractor.Items[CoB_contractor.SelectedIndex].ToString().Replace("\r", "")}\" ");

                        sb.Append($"\"{CB_addSignature.IsChecked.Value.ToString()}\" ");
                        sb.Append($"\"{CB_addStamp.IsChecked.Value.ToString()}\" ");
                        sb.Append($"\"{CB_sendEmail.IsChecked.Value.ToString()}\" ");
                        break;
                    case ScriptQueryType.GET_CONTRACTORS:
                        sb.Append($" \"{TB_dbPath.Text}\" ");
                        break;
                    case ScriptQueryType.GET_CONTRACTOR_DATES:
                        if ((string)CoB_contractor.Items[CoB_contractor.SelectedIndex] == "Все контрагенты")
                            throw new Exception($"Выберите контрагента");
                        sb.Append($" \"{TB_dbPath.Text}\" ");
                        sb.Append($"\"{CoB_contractor.Items[CoB_contractor.SelectedIndex].ToString().Replace("\r", "")}\" ");
                        break;
                }
            });
            return sb.ToString();
        }

        private void CoB_contractorSelectionChange(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                if (CB_filterContractor.IsChecked == true)
                    _app.StartScript(ScriptQueryType.GET_CONTRACTOR_DATES);
            });
        }

        private void CB_Change(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox checkBox)
            {
                switch (checkBox.Name)
                {
                    case "CB_filterPeriod":
                        DP_startDate.IsEnabled = checkBox.IsChecked.Value;
                        DP_endDate.IsEnabled = checkBox.IsChecked.Value;
                        break;
                    case "CB_filterContractor":
                        CoB_contractor.IsEnabled = checkBox.IsChecked.Value;
                        break;
                }
            }
        }

        private void BT_folder_overiew_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog
            {
                IsFolderPicker = true,
                Title = "Выбор папки сохранения",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
                TB_outputPath.Text = dlg.FileName;      
        }

        private void BT_overiew_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is TextBox targetTextBox)
            {
                string fileFilter;
                string dialogTitle;
                switch(button.Name)
                {
                    case "BT_dbPathOveriew":
                        fileFilter = "Базы данных (*.xlsx)|*.xlsx";
                        dialogTitle = "Открыть базу данных";
                        break;
                    case "BT_docTemplateOveriew":
                    case "BT_mailTemplateOveriew":
                        fileFilter = "Шаблоны документов (*.docx, *.dotx)|*.docx;*.dotx";
                        dialogTitle = "Открыть шаблон документа";
                        break;
                    case "BT_signaturePathOveriew":
                    case "BT_stampPathOveriew":
                        fileFilter = "Изображения с поддержкой прозрачности (*.png)|*.png";
                        dialogTitle = "Открыть изображение";
                        break;
                    default:
                        return;
                }

                var dialog = new OpenFileDialog
                {
                    Title = dialogTitle,
                    Filter = fileFilter,
                    CheckFileExists = true,
                    CheckPathExists = true,
                    Multiselect = false
                };

                bool? result = dialog.ShowDialog();

                if (result == true)
                    targetTextBox.Text = dialog.FileName;

                if (button.Name == "BT_dbPathOveriew")
                    _app.StartScript(ScriptQueryType.GET_CONTRACTORS);
            }
        }

        private void BT_startGenerating_Click(object sender, RoutedEventArgs e)
        {
            _app.StartScript(ScriptQueryType.GENERATE);
        }
    }
}
