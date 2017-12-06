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
        string date;
        string time, shitToPrint;
        string myFileName, filename, filenameDOC;
        string[] words;
        string tokens;

        public MainWindow()
        {
            date = now.ToString("dd-MM-yyyy");
            time = now.ToString("HH-mm-ss");
            myFileName = $"{date}-{time}.xls";

            InitializeComponent();
        }


        //Generate Excel file
        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {

            // TODO: verify if user opened a file
            if (filenameDOC != null)
            {

                // TODO: verify if user checked at least one box
                if(NameCheckBox.IsChecked == false && PeriodCheckBox.IsChecked == false && IDCheckBox.IsChecked == false && StatusCheckBox.IsChecked == false && GPACheckBox.IsChecked == false && FailCheckBox.IsChecked == false)
                {
                    MessageBox.Show("Debes elegir al menos una opcion", "¡Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    // TODO: write or append to xls file
                    try
                    {
                        tokens = File.ReadAllText($"{filenameDOC}.doc");
                        words = tokens.Split(' ');

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Hubo un error al intentar abrir el archivo\n{ex}", "¡Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    
                    /*for (int i = 0; i < words.Length; i++)
                    {
                        Console.WriteLine($"{words[i]} {i}");
                    }*/


                    Debug.WriteLine(myFileName);

                    if (NameCheckBox.IsChecked == true)
                    {
                        for (int i = 0; i < words.Length; i++)
                        {
                            switch(i)
                                {
                                case 70:

                                    shitToPrint += $"{words[i]} ";
                                    break;


                                case 71:

                                    shitToPrint += $"{words[i]} ";
                                    break;

                                case 72:

                                    shitToPrint += $"{words[i]} ";
                                    break;

                                case 73:

                                    shitToPrint += $"{words[i]}\t";
                                    break;
                                    
                            }
                        }
                    }

                    // If the file doesn't exists
                    if (!File.Exists(myFileName))
                    {
                        using (StreamWriter sw = File.CreateText(myFileName))
                        {
                            sw.WriteLine(shitToPrint);
                            MessageBox.Show("Archivo creado con exito y alumno agregado!", "Correcto!", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    else
                    {
                        // File already exists
                        using (StreamWriter sw = File.AppendText(myFileName))
                        {
                            sw.WriteLine(shitToPrint);
                            MessageBox.Show("Alumno agregado!", "Correcto!", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Debe primero seleccionar un cardex!", "¡Error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        //Open pdf file
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
                            filename = Path.GetFullPath(openFileDialog.FileName);
                            filenameDOC = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                            MessageBox.Show($"Abriendo archivo {filename}", "Informacion", MessageBoxButton.OK, MessageBoxImage.Information);
                            PDFParser pdfParser = new PDFParser();
                            // extract the text from a pdf file and exports it to a doc file
                            pdfParser.ExtractText(Path.GetFullPath(openFileDialog.FileName), Path.GetFileNameWithoutExtension(filename)+".doc");
                        }
                    }
                }
                catch (InvalidOperationException)
                {
                    //User cancels file choosing
                }
            }
        }
    }
}
