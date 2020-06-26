using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

/////Add the references (new)
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Drawing.Drawing2D;
////

namespace Picojazz_Constat_V1._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //Methode Find and Replace:
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText,
                        ref matchCase, ref matchWholeWord,
                        ref matchWildCards, ref matchSoundLike,
                        ref nmatchAllForms, ref forward,
                        ref wrap, ref format, ref replaceWithText,
                        ref replace, ref matchKashida,
                        ref matchDiactitics, ref matchAlefHamza,
                        ref matchControl);
        }



        //Methode Create the document :
        private void CreateWordDocument(object filename, object savaAs)
        {
            List<int> processesbeforegen = getRunningProcesses();
            object missing = Missing.Value;
            string tempPath = null;

            Word.Application wordApp = new Word.Application();

            Word.Document aDoc = null;

            if (File.Exists((string)filename))
            {
                DateTime today = DateTime.Now;

                object readOnly = false; //default
                object isVisible = false;

                wordApp.Visible = false;

                aDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing);

                aDoc.Activate();

                //Find and replace:
                this.FindAndReplace(wordApp, "$dateinf$", dateInfoTextBox.Text);
                this.FindAndReplace(wordApp, "$clerk$", nomClerkTextBox.Text);
                this.FindAndReplace(wordApp, "$heure$", heureTextBox.Text);
                this.FindAndReplace(wordApp, "$lieu$", lieuTextBox.Text);

                //vehicule 1
                this.FindAndReplace(wordApp, "$imma$", immaTextBox.Text);
                this.FindAndReplace(wordApp, "$marque$", marqueTextBox.Text);
                this.FindAndReplace(wordApp, "$type$", typeTextBox.Text);
                this.FindAndReplace(wordApp, "$genre$", genreTextBox.Text);
                this.FindAndReplace(wordApp, "$misacir$", misaTextBox.Text);
                this.FindAndReplace(wordApp, "$vistec$", visTecTextBox.Text);
                this.FindAndReplace(wordApp, "$ass$", ass.Text);
                this.FindAndReplace(wordApp, "$police$", policeTextBox.Text);
                this.FindAndReplace(wordApp, "$valvec$", valAssTextBox.Text);
                this.FindAndReplace(wordApp, "$auvec$", auAssTextBox.Text);
                this.FindAndReplace(wordApp, "$nomciv$", nomCivTextBox.Text);
                this.FindAndReplace(wordApp, "$adrciv$", adrCivTextBox.Text);
                this.FindAndReplace(wordApp, "$nomcon$", nomConTextBox.Text);
                this.FindAndReplace(wordApp, "$adrcon$", adrConTextBox.Text);
                this.FindAndReplace(wordApp, "$datenaiss$", datenaiss.Text);
                this.FindAndReplace(wordApp, "$perm$", permTextBox.Text);
                this.FindAndReplace(wordApp, "$valperm$", valPermTextBox.Text);
                this.FindAndReplace(wordApp, "$auperm$", auPermTextBox.Text);
                this.FindAndReplace(wordApp, "$deg$", deg1TextBox.Text);
                this.FindAndReplace(wordApp, "$decla$", decla.Text);


                //vehicule 2

                this.FindAndReplace(wordApp, "$imma2$", imma2.Text);
                this.FindAndReplace(wordApp, "$marque2$", mar2.Text);
                this.FindAndReplace(wordApp, "$type2$", typeVec2.Text);
                this.FindAndReplace(wordApp, "$genre2$", genreVec2.Text);
                this.FindAndReplace(wordApp, "$miscir2$", misacir2.Text);
                this.FindAndReplace(wordApp, "$vistec2$", vistec2.Text);
                this.FindAndReplace(wordApp, "$ass2$", ass2.Text);
                this.FindAndReplace(wordApp, "$police2$", police2.Text);
                this.FindAndReplace(wordApp, "$valvec2$", valass2.Text);
                this.FindAndReplace(wordApp, "$auvec2$", auass2.Text);
                this.FindAndReplace(wordApp, "$nomciv2$", nomciv2.Text);
                this.FindAndReplace(wordApp, "$adrciv2$", adrciv2.Text);
                this.FindAndReplace(wordApp, "$nomcon2$", nomcon2.Text);
                this.FindAndReplace(wordApp, "$adrcon2$", adrcon2.Text);
                this.FindAndReplace(wordApp, "$datenaiss2$", datenaiss2.Text);
                this.FindAndReplace(wordApp, "$perm2$", perm2.Text);
                this.FindAndReplace(wordApp, "$valperm2$", valperm2.Text);
                this.FindAndReplace(wordApp, "$auperm2$", auperm2.Text);
                this.FindAndReplace(wordApp, "$deg2$", deg2.Text);
                this.FindAndReplace(wordApp, "$decla2$", decla2.Text);






            }
            else
            {
                MessageBox.Show("fichier non crée.");
                return;
            }

            //Save as: filename
            aDoc.SaveAs2(ref savaAs, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            //Close Document:
            //aDoc.Close(ref missing, ref missing, ref missing);
            //File.Delete(tempPath);
            MessageBox.Show("fichier crée.");
            List<int> processesaftergen = getRunningProcesses();
            killProcesses(processesbeforegen, processesaftergen);
        }


        public List<int> getRunningProcesses()
        {
            List<int> ProcessIDs = new List<int>();
            //here we're going to get a list of all running processes on
            //the computer
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains("WINWORD"))
                {
                    ProcessIDs.Add(clsProcess.Id);
                }
            }
            return ProcessIDs;
        }

        private void killProcesses(List<int> processesbeforegen, List<int> processesaftergen)
        {
            foreach (int pidafter in processesaftergen)
            {
                bool processfound = false;
                foreach (int pidbefore in processesbeforegen)
                {
                    if (pidafter == pidbefore)
                    {
                        processfound = true;
                    }
                }

                if (processfound == false)
                {
                    Process clsProcess = Process.GetProcessById(pidafter);
                    clsProcess.Kill();
                }
            }
        }





        private void recupFichierModel_Click(object sender, EventArgs e)
        {
            if (loadFichier.ShowDialog() == DialogResult.OK)
            {
                fichierTextBox.Text = loadFichier.FileName;
                
            }
        }

        

        private void saveBtn_Click_1(object sender, EventArgs e)
        {
            if (saveDoc.ShowDialog() == DialogResult.OK)
            {
                CreateWordDocument(fichierTextBox.Text, saveDoc.FileName);
                
               printDocument1.DocumentName = saveDoc.FileName;
            }
        }

        private void fichierTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void dateInfoTextBox_TextChanged(object sender, EventArgs e)
        {
            dateInfoTextBox.Text = this.dateInfoTextBox.Text.ToUpper();
            this.dateInfoTextBox.SelectionStart = this.dateInfoTextBox.Text.Length;
        }

        private void lieuTextBox_TextChanged(object sender, EventArgs e)
        {
            lieuTextBox.Text = this.lieuTextBox.Text.ToUpper();
            this.lieuTextBox.SelectionStart = this.lieuTextBox.Text.Length;
        }
    }
}
