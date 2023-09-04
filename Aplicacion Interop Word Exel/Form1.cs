using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Security.Cryptography.X509Certificates;
using System.Drawing.Text;
using System.IO;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.Remoting.Contexts;

namespace Aplicacion_Interop_Word_Exel
{
    public partial class Form1 : Form
    {
        private string fileName;
        private dynamic text;
        private object ruta;

        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialogo = new SaveFileDialog();
            if (dialogo.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string ruta = dialogo.FileName;
            var WordApp = new word.Application();
            WordApp.Visible = true;
            WordApp.Documents.Add();
            String Texto = textBox1.Text;
            WordApp.Selection.TypeText(Texto);
            WordApp.ActiveDocument.SaveAs2(ruta);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            SaveFileDialog dialogo = new SaveFileDialog();
            if (!Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "EE_Commande_Fournisseur"))
            {
                Directory.CreateDirectory(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\EE_Commande_Fournisseur");
            }
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Add();
            String Texto = textBox1.Text;
            excelApp.Selection.TypeText(Texto);
            excelApp.ActiveSheet.SaveAs2(ruta);
            var worksheet = (Excel.Worksheet)workbook.Worksheets[1];
            excelApp.Visible = true;
            worksheet.Cells[1, 1] = text;
           



        }
    }
    }

    

