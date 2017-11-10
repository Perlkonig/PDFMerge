using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows;
using swf = System.Windows.Forms;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml;
using System.Collections;
using System.Data;
using System.Data.OleDb;

namespace PDFMerge
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

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void txt_Excel_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Note that you can have more than one file.
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // Assuming you have one file that you care about, pass it off to whatever
                // handling code you have defined.
                txt_Excel.Text = files[0];
            }
        }

        private void txt_Excel_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void txt_PDF_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void txt_Output_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void txt_PDF_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Note that you can have more than one file.
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // Assuming you have one file that you care about, pass it off to whatever
                // handling code you have defined.
                txt_PDF.Text = files[0];
            }
        }

        private void txt_Output_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Note that you can have more than one file.
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // Assuming you have one file that you care about, pass it off to whatever
                // handling code you have defined.
                txt_Output.Text = files[0];
            }
        }

        private void btn_Excel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "Excel files (*.xls; *.xlsx; *.xlsm)|*.xls;*.xlsx;*.xlsm|All files (*.*)|*.*";
            if (d.ShowDialog() == true)
            {
                txt_Excel.Text = d.FileName;
            }
        }

        private void btn_PDF_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*";
            if (d.ShowDialog() == true)
            {
                txt_PDF.Text = d.FileName;
            }
        }

        private void btn_Output_Click(object sender, RoutedEventArgs e)
        {
            swf.FolderBrowserDialog d = new swf.FolderBrowserDialog();
            swf.DialogResult result = d.ShowDialog();
            if (result == swf.DialogResult.OK)
            {
                txt_Output.Text = d.SelectedPath;
            }
        }

        private List<string> GetFieldNames()
        {
            List<string> flist = new List<string>();
            PdfReader pdfReader = new PdfReader(txt_PDF.Text);

            var fields = pdfReader.AcroFields.Fields;
            foreach (var key in fields.Keys)
            {
                flist.Add(key);
            }
            return flist;
        }

        private List<Dictionary<string,string>> LoadExcel()
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook wb = app.Workbooks.Open(txt_Excel.Text);
            Excel.Worksheet ws = wb.Worksheets["Sheet1"];
            Excel.Range usedRange = ws.UsedRange;
            List<Dictionary<string, string>> l = new List<Dictionary<string, string>>();

            //Iterate the rows in the used range
            Excel.Range firstrow = null;
            foreach (Excel.Range row in usedRange.Rows)
            {
                if (firstrow == null)
                {
                    firstrow = row;
                } else
                {
                    Dictionary<string, string> d = new Dictionary<string, string>();
                    for (int i = 0; i < row.Columns.Count; i++)
                        d.Add(firstrow.Cells[1,i+1].Value2.ToString(), row.Cells[1, i + 1].Value2.ToString());
                    l.Add(d);
                }
            }
            return l;
        }

        private void btn_Run_Click(object sender, RoutedEventArgs e)
        {
            var fieldNames = GetFieldNames();
            var data = LoadExcel();

            int count = 0;
            string outdir = txt_Output.Text;
            if (outdir == String.Empty)
            {
                outdir = ".";
            }

            foreach (var row in data)
            {
                count++;
                string outpath = outdir + @"\_merged" + count.ToString() + ".pdf";
                using (FileStream outFile = new FileStream(outpath, FileMode.Create))
                {
                    PdfReader source = new PdfReader(txt_PDF.Text);
                    PdfStamper copy = new PdfStamper(source, outFile);
                    AcroFields fields = copy.AcroFields;

                    foreach (var field in fieldNames)
                    {
                        fields.SetField(field, row[field].ToString());
                    }

                    copy.Close();
                    source.Close();
                }

            }

            MessageBox.Show("Merge complete.");
        }
    }
}
