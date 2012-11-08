using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace DocumentEditorTestApp
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

        private void menuItemOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if(dialog.ShowDialog() == true)
            {
                //docEditor.OpenDocxFile(dialog.FileName);
                FlowDocument fd = new FlowDocument();
                FileStream fileStream = File.OpenRead(dialog.FileName);
                fd.LoadOpenXml(fileStream);
                docEditor.rtbDocument.Document = fd;
            }
        }

        private void menuItemSaveFile_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            if(dialog.ShowDialog() == true)
            {
                FileStream saveFileStream = dialog.OpenFile() as FileStream;
                //docEditor.ConvertToOpenXml(saveFileStream);
                docEditor.rtbDocument.Document.ConvertToOpenXml(saveFileStream);
            }
        }
    }
}
