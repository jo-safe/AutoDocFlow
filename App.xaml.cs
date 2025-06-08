using System;
using System.IO;
using System.Windows;
using System.Diagnostics;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Text.Json;
using System.Linq;

namespace AutoDocFlow
{
	public enum ScriptQueryType
	{
		GENERATE = 0,
		GET_CONTRACTORS = 1,
		GET_CONTRACTOR_DATES = 2
	}

	public partial class App : Application
    {
		MainWindow _mainWindow;

        static readonly string settingsPath = "settings.cfg";
		public List<string> contractorsList = new List<string>();

		protected override void OnStartup(StartupEventArgs e)
		{
			base.OnStartup(e);

            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

            _mainWindow = new MainWindow(this);
            _mainWindow.Show();
            LoadSettings();

            this.Exit += App_Exit;
        }

        private void App_Exit(object sender, ExitEventArgs e)
        {
			SaveSettings();
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            MessageBox.Show("UI exception caught:\n" + e.Exception.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            MessageBox.Show("Non-UI thread exception:\n" + ex?.Message, "Fatal Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void TaskScheduler_UnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            MessageBox.Show("Task exception:\n" + e.Exception.Message, "Task Error", MessageBoxButton.OK, MessageBoxImage.Error);
            e.SetObserved();
        }

        private void LoadSettings()
		{
            try
            {
                string settingsPath = "settings.cfg";
                
                var lines = File.ReadAllLines(settingsPath);
                if (lines.Length < 8) return;

                _mainWindow.UpdatePaths(File.Exists(lines[0]) ? lines[0] : "",
                                                        File.Exists(lines[0]) ? lines[1] : "",
                                                        File.Exists(lines[0]) ? lines[2] : "",
                                                        File.Exists(lines[0]) ? lines[3] : "",
                                                        File.Exists(lines[0]) ? lines[4] : "",
                                                        File.Exists(lines[0]) ? lines[5] : "",
                                                        File.Exists(lines[0]) ? lines[6] : "",
                                                        File.Exists(lines[0]) ? lines[7] : "",
                                                        File.Exists(lines[0]) ? lines[8] : "",
                                                        File.Exists(lines[0]) ? lines[9] : "");

                if (File.Exists(lines[0]))
                    StartScript(ScriptQueryType.GET_CONTRACTORS);
            }
            catch (Exception e)
            {
                MessageBox.Show("Файл настроек отсутствует или поврежден.");
            }
		}

        public void SaveSettings()
        {
            try
            {
				List<string> settings = _mainWindow.GetSettings();
                File.WriteAllLines(settingsPath, settings);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении настроек: {ex.Message}");
            }
        }

        public async void StartScript(ScriptQueryType queryType)
		{
			try
			{
				string arguments = _mainWindow.GetScriptParams(queryType);

                string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "script.exe");

                if (!File.Exists(scriptPath))
                {
                    MessageBox.Show("Скрипт не найден.");
                    return;
                }

                ProcessStartInfo psi = new ProcessStartInfo
                {
                	FileName = scriptPath,
                	Arguments = $"{arguments}",
                	UseShellExecute = false,
                	RedirectStandardOutput = true,
                	RedirectStandardError = true,
                	CreateNoWindow = true,
                	WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
                };

                var process = new Process { StartInfo = psi };
                process.Start();

                string output = process.StandardOutput.ReadToEnd();
                if (output.Length >= 2)
                    output = output.Substring(0, output.Length - 2);
                string error = process.StandardError.ReadToEnd();
                if (error.Contains("Traceback")) 
                {
                    MessageBox.Show("Ошибка выполнения: " + error);
                    return;
                }
                process.WaitForExit();
                
                switch (queryType)
                {
                    case ScriptQueryType.GENERATE:
                        _mainWindow.TB_log.AppendText(output);
                        break;
                    case ScriptQueryType.GET_CONTRACTORS:
                        List<string> contractors = output.Split("\r\n").ToList();
                        contractors.Insert(0, "Все контрагенты");
                        _mainWindow.UpdateContractorsList(contractors.Take(contractors.Count - 1).ToList());
                        break;
                    case ScriptQueryType.GET_CONTRACTOR_DATES:
                        string[] dates = output.Split(' ');
                        _mainWindow.UpdateDates(DateTime.ParseExact(dates[0], "dd.MM.yyyy", null), DateTime.ParseExact(dates[1], "dd.MM.yyyy", null));
                        break;
                }
                }
			catch (Exception ex)
			{
				MessageBox.Show($"Ошибка выполнения: {ex.Message}");
			}
		}
	}
}
