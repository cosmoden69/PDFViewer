using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace PDFViewer
{
    public partial class Form1 : Form
    {
        string docFile = "";
        string pdfFile = "";

        public Form1(string[] args)
        {
            docFile = args[0];
            this.Text = "보고서조회" + (args.Length > 1 ? " - " + args[1] : "");

            InitializeComponent();

            this.Load += Form1_Load;
            this.Disposed += Form1_Disposed;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Location = new Point(Location.X, 0);
            Size = new Size(Width, Screen.PrimaryScreen.WorkingArea.Height);
            LoadDocument();
        }

        private void Form1_Disposed(object sender, EventArgs e)
        {
            File.Delete(pdfFile);
        }

        private void LoadDocument()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                pdfFile = Path.ChangeExtension(docFile, "pdf");
                if (File.Exists(pdfFile)) File.Delete(pdfFile);
                convertDOCtoPDF(docFile, pdfFile);
                //using (FileStream fs = new FileStream(pdfFile, FileMode.Open, FileAccess.Read))
                //{
                //    pdf.LoadDocument(fs);
                //}
                pdf.LoadDocument(pdfFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void convertDOCtoPDF(string docfile, string pdffile)
        {
            object misValue = System.Reflection.Missing.Value;
            String PATH_APP_PDF = pdffile;

            var WORD = new Word.Application();

            Word.Document doc = WORD.Documents.Open(docfile);
            doc.Activate();

            doc.SaveAs2(@PATH_APP_PDF, Word.WdSaveFormat.wdFormatPDF, misValue, misValue, misValue,
            misValue, misValue, misValue, misValue, misValue, misValue, misValue);

            doc.Close();
            WORD.Quit();

            releaseObject(doc);
            releaseObject(WORD);

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                //TODO
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
