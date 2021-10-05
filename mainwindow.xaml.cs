using System;
using System.IO;
using System.Linq.Expressions;
using System.Threading.Tasks;
using System.Threading;
using System.Windows;


namespace ReferenceKeyTesting_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            Main = this;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
        }
        internal static MainWindow Main;
        
        internal string UpdateLabel
        {
            set { Dispatcher.Invoke(new Action(() => { LblNoOfKeysCollected.Content = value; })); }
        }
        private InventorInteraction _inventorInteraction;
        public static CancellationTokenSource Cts;
        public static CancellationToken Token;
        private int _count;
        private bool _isRunning = false;
        private async void BtnGenKeyContextAtLast_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning) { MessageBox.Show("Click 'Stop Collection' Button First"); return; }
            _isRunning = true;
            _inventorInteraction = new InventorInteraction();
            Cts = new CancellationTokenSource();
            Token = Cts.Token;
            BtnCollectContextKey.Visibility = Visibility.Visible;
            LblContextKeys.Visibility = Visibility.Visible;
            _count = 0;
            LblContextKeys.Content = _count.ToString();
            try
            {
                await Task.Run(() => _inventorInteraction.CollectDataForActiveAssembly());
            }
            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
            finally
            {
                _isRunning = false;
                Cts.Dispose();
            }
        }
        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Cts == null) MessageBox.Show("Click on Generate button first!");
                else
                {
                    Cts.Cancel();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Operation Stopped!");
            }
        }
        private async void BtnGenKeyContextEveryTime_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning) { MessageBox.Show("Click 'Stop Collection' Button First"); return; }
            _isRunning = true;
            _inventorInteraction = new InventorInteraction();
            Cts = new CancellationTokenSource();
            Token = Cts.Token;
            LblContextKeys.Visibility = Visibility.Hidden;

            BtnCollectContextKey.Visibility = Visibility.Hidden;
            try
            {
                await Task.Run(() => _inventorInteraction.CollectDataForActiveAssembly(false));
            }
            catch (Exception ex)
            {
                Extension.CreateLog(ex);
            }
            finally
            {
                _isRunning = false;
                Cts.Dispose();
            }
        }
        private void BtnOpenExcel_Click(object sender, RoutedEventArgs e)
        {
            ExcelInteraction excelInteraction = new ExcelInteraction();
        }

        private void BtnGoToTester_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning) { MessageBox.Show("Click 'Stop Collection' Button First"); return; }
            Reference_Key_Tester tester = new Reference_Key_Tester();
            tester.Show();
            this.Close();
        }
        private void BtnLogFile_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(Environment.CurrentDirectory + @"\log.txt");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_isRunning)
            {
                Cts.Dispose();
                _inventorInteraction.Close();
            }
        }

        private void BtnClearLog_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Clear the Log?", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
            if (result == MessageBoxResult.Yes) Extension.ClearLog();
        }

        private void BtnClearExcel_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show("Clear all the Keys?", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
            if (result == MessageBoxResult.Yes) {Extension.ClearExcel(); LblNoOfKeysCollected.Content = "0";}
        }

        private void BtnCollectContextKey_Click(object sender, RoutedEventArgs e)
        {
            if (_isRunning)
            {
                _count++;
                _inventorInteraction.SaveKeyContextOnClick();
                LblContextKeys.Content = _count.ToString();
            }
            else
            {
                MessageBox.Show("Click on Generate Button First!");
            }
        }
    }
}
