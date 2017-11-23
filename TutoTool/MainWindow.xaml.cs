using Microsoft.Win32;
using System;
using System.IO;
using System.Windows;
using System.Diagnostics;

namespace TutoTool
{
    public partial class MainWindow : Window
    {
        // I know this is shit but i can't think of any other solution
        // Create a file in which this sesion will work
        DateTime now = DateTime.Now;
        int count = 0;
        string date;
        string time;
        string myFileName;


        // LEARNING HOW TO TELL TIME


        public MainWindow()
        {


            date = now.ToString("dd-MM-yyyy");
            time = now.ToString("HH-mm-ss");
            myFileName = $"{date}-{time}.xls";

            InitializeComponent();
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            // TODO: verify if user checked at least one box

            // TODO: verify if user opened a file

            // TODO: write or append to xls file

            count++;
            Debug.WriteLine(myFileName);
            string shitToPrint = $"Test # {count} \t check \t it \t out";

            // If the file doesn't exists
            if (!File.Exists(myFileName))
            {
                using (StreamWriter sw = File.CreateText(myFileName))
                {
                    sw.WriteLine(shitToPrint);
                }
            }
            else
            {
                // File already exists
                using (StreamWriter sw = File.AppendText(myFileName))
                {
                    sw.WriteLine(shitToPrint);
                }
            }
        }

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {

            Stream myStream = null;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Selecciona el archivo",
                Filter = "Archivos PDF (*.pdf)|*.pdf",
                FilterIndex = 2,
                RestoreDirectory = true
            };
            bool? result = openFileDialog.ShowDialog();

            if (result != null)
            {
                try
                {
                    if ((myStream = openFileDialog.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            string filename = Path.GetFileName(openFileDialog.FileName);
                            MessageBox.Show($"Abriendo archivo {filename}", "Informacion", MessageBoxButton.OK, MessageBoxImage.Information);
                        }

                        // Reading the opened file
                    }
                }
                catch (InvalidOperationException)
                {
                    // Users cancels
                }
            }
        }
    }
}
