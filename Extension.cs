using System;
using System.IO;
using System.Windows;

namespace ReferenceKeyTesting_WPF
{
    public static class Extension
    {
        private static readonly string Path = Environment.CurrentDirectory + @"\log.txt";
        public static bool IsFileOpen(string filePath)
        {
            bool rtnValue = false;
            try
            {
                FileStream fs = File.OpenWrite(filePath);
                fs.Close();
            }
            catch (Exception)
            {
                rtnValue = true;
            }
            return rtnValue;
        }

        public static void CreateLog(Exception ex)
        {
            string timeStamp = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            if (ex is OutOfMemoryException)
            {
                MessageBox.Show(timeStamp + Environment.NewLine + "Out Of Memory Exception Thrown!", "Warning", MessageBoxButton.OK);
                MainWindow.Cts.Cancel();
            }
            File.AppendAllText(Path, Environment.NewLine + Environment.NewLine + timeStamp + Environment.NewLine + ex.StackTrace + Environment.NewLine + ex.Message);


        }

        public static void ClearLog()
        {
            File.WriteAllText(Path, String.Empty);
        }

        public static void ClearExcel()
        {
            try
            {
                File.Delete(ExcelInteraction.Path);
            }
            catch (Exception e)
            {
                MessageBox.Show("Close the Excel!");
            }

        }

    }
}
