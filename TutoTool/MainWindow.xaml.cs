using Microsoft.Win32;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace TutoTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("En construcción");
        }

        private void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Selecciona el archivo",
                Filter = "Archivos de Texto (*.txt)|*.txt",
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
                            MessageBox.Show("Ya se abrio el archivo");
                        }

                        // Reading the opened file
                    }
                }
                catch (InvalidOperationException)
                {
                    // Do nothing if user cancels
                }
            }
        }
    }
}
