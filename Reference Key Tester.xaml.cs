using System;
using System.Collections.Generic;
using System.IO;
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

namespace ReferenceKeyTesting_WPF
{
    /// <summary>
    /// Interaction logic for Reference_Key_Tester.xaml
    /// </summary>
    public partial class Reference_Key_Tester : Window
    {
        public Reference_Key_Tester()
        {
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            InitializeComponent();
        }

        private void BtnGoBack_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void BtnOpenExcel_Click(object sender, RoutedEventArgs e)
        {
            ExcelInteraction excelInteraction = new ExcelInteraction();
        }

        private void BtnBind_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(TxtKeyContext.Text)  && !string.IsNullOrEmpty(TxtReferenceKey.Text))
            {
                InventorInteraction inventorInteraction = new InventorInteraction();
               bool flag= inventorInteraction.CheckReferenceKey(TxtReferenceKey.Text, TxtKeyContext.Text);
               TxtResult.Text = flag ? "Reference key binds back to the entity successfully!" : "Reference key failed to bind back to the entity!";
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
        }

        private void TxtReferenceKey_TextChanged(object sender, TextChangedEventArgs e)
        {
            TxtResult.Text = "";
        }

        private void TxtKeyContext_TextChanged(object sender, TextChangedEventArgs e)
        {
            TxtResult.Text = "";
        }
    }
}
