using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DopisiGenerator
{
    using word = Microsoft.Office.Interop.Word;
    public partial class GlavnaForma : Form
    {
        word.Application app;
        Document doc;

        public GlavnaForma()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            app = new word.Application();
            using (StreamReader sr = new StreamReader(openCompanyDatabaseFile.FileName))
            {
                string line;
                
                while ((line = sr.ReadLine()) != null)
                {
                    doc = app.Documents.Open(openTemplateDialog.FileName);
                    //imeFirme,ime,telefon,email
                    var data = line.Split(',');
                    app.Selection.Find.Execute(tbCompanyName.Text, true, true, false, false, false, true, 1, false, data[0], 2);
                    app.Selection.Find.Execute(tbName.Text, true, true, false, false, false, true, 1, false, data[1], 2);
                    app.Selection.Find.Execute(tbPhone.Text, true, true, false, false, false, true, 1, false, data[2], 2);
                    app.Selection.Find.Execute(tbEmail.Text, true, true, false, false, false, true, 1, false, data[3], 2);
                    var folderPath = System.IO.Path.GetDirectoryName(openTemplateDialog.FileName);
                    doc.SaveAs2(folderPath + "\\Dopis " + data[0]+".pdf", word.WdSaveFormat.wdFormatPDF);
                    doc.Close(false);
                }
                app.Quit();
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            openTemplateDialog.Filter =
               "Word files (*.docx)|*.docx";
            openTemplateDialog.Title = "Select the template file";
            openTemplateDialog.ShowDialog();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            openCompanyDatabaseFile.ShowDialog();
        }
    }
}
