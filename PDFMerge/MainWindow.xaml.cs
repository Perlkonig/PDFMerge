using System;
using System.Collections.Generic;
using System.Linq;
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
            string pdfTemplate = txt_PDF.Text;

            // create a new PDF reader based on the PDF template document
            PdfReader pdfReader = new PdfReader(pdfTemplate);

            // create and populate a string builder with each of the
            // field names available in the subject PDF
            StringBuilder sb = new StringBuilder();
            sb.Append("Found the following fields:" + Environment.NewLine);

            var fields = pdfReader.AcroFields.Fields;
            foreach (var key in fields.Keys)
            {
                flist.Add(key);
            }
            return flist;
        }

        private List<Dictionary<string,string>> LoadExcel()
        {
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txt_Excel.Text + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO\"");
            con.Open();

            string sheetName = "Sheet1$"; // (string)dt.Rows[0]["Table_Name"];

            DataTable xlWorksheet = new DataTable();

            xlWorksheet.Load(new OleDbCommand("Select * From ["+ sheetName +"]", con).ExecuteReader());
            List<Dictionary<string, string>> l = new List<Dictionary<string, string>>();

            for (int nRow = 1; nRow < xlWorksheet.Rows.Count; nRow++)
            {
                Dictionary<string, string> d = new Dictionary<string, string>();

                for (int nColumn = 0; nColumn < xlWorksheet.Columns.Count; nColumn++)
                {
                    d.Add(xlWorksheet.Rows[0].ItemArray[nColumn].ToString(), xlWorksheet.Rows[nRow].ItemArray[nColumn].ToString());
                }

                l.Add(d);
            }
            return l;
        }

        private void btn_Run_Click(object sender, RoutedEventArgs e)
        {
            var fields = GetFieldNames();
            var data = LoadExcel();
            StringBuilder sb = new StringBuilder();
            foreach (var row in data)
            {
                foreach (var field in fields)
                {
                    sb.Append(row[field].ToString() + ", ");
                }
                sb.Append(Environment.NewLine);
            }
            MessageBox.Show(sb.ToString());
        }
    }
}
