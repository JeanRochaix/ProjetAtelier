using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Atelier_GenerationExercideVerbe_Jean
{
    public partial class Form1 : Form
    {
        #region Déclarations variables
        Random r = new Random();
        //Listes
        //List<PositionVerbe> lstxxx;
        List<object> lstFields;
        List<object> lstEndFields;
        List<object> lstPosStart;
        List<object> lstPosEnd;
        List<object> lstFieldsVerd;
        List<object> lstFieldsVerf;
        List<object> lstPosStartV;
        List<object> lstPosEndV;
        List<object> lstFieldsImgD;
        List<object> lstFieldsImgF;
        List<object> lstPosStartI;
        List<object> lstPosEndI;
        Microsoft.Office.Interop.Word.Document nvDoc;
        Microsoft.Office.Interop.Word.Application msWord;
        object missing;
        Random rnd;
       // int irnd;
        #endregion
        public Form1()
        {
            InitializeComponent();
            rnd = new Random();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // Connexion à Word
            msWord = new Microsoft.Office.Interop.Word.Application();
            msWord.Visible = false; // mettez cette variable à true si vous souhaitez visualiser les opérations.
            missing = System.Reflection.Missing.Value;

            // Attribuer le nom
            object fileName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Fiche_P5.docx";

            //Microsoft.Office.Interop.Word.Document nvDoc;

            // Tester s'il s'agit d'un nouveau document ou d'un document existant.
            if (System.IO.File.Exists((string)fileName))
            {
                // ouvrir le document existant
                nvDoc = msWord.Documents.Open(ref fileName, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing);
            }
            else
            {
                // Choisir le template
                object templateName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Template\Test.dotm";
                // Créer le document
                nvDoc = msWord.Documents.Add(ref templateName, ref missing, ref missing, ref missing);
            }

            lstFieldsVerf = new List<object> { };
            lstFields = new List<object> { };
            lstEndFields = new List<object> { };
            lstPosStart = new List<object> { };
            lstPosEnd = new List<object> { };
            lstPosStartV = new List<object> { };
            lstPosEndV = new List<object> { };
            lstFieldsVerd = new List<object> { };
            lstFieldsImgD = new List<object> { };
            lstFieldsImgF = new List<object> { };
            lstPosStartI = new List<object> { };
            lstPosEndI = new List<object> { };
            //boucle for qui rempli ligne par ligne le document 
            for (int i = 1; i <= 10 ; i++)
            {
                lstFieldsVerd.Add("dverbe" + i);
                lstFieldsVerf.Add("fverbe" + i);

                lstEndFields.Add("fpronom" + i);
                lstFields.Add("dpronom" + i);

                lstFieldsImgD.Add("imaged" + i);
                lstFieldsImgF.Add("imagef" + i);

                lstPosStart.Add(new Object());
                lstPosStart[lstPosStart.Count - 1] = nvDoc.Bookmarks.get_Item(lstFields[i - 1]).Start;
                lstPosEnd.Add(new Object());
                lstPosEnd[lstPosEnd.Count - 1] = nvDoc.Bookmarks.get_Item(lstEndFields[i - 1]).End;

                lstPosStartV.Add(new Object());
                lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i - 1]).Start;
                lstPosEndV.Add(new Object());
                lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i - 1]).End;

                lstPosStartI.Add(new Object());
                lstPosStartI[lstPosStartI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgD[i - 1]).Start;
                lstPosEndI.Add(new Object());
                lstPosEndI[lstPosEndI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgF[i - 1]).End;

                PronomText(i - 1);
                VerbeTexte(i - 1);
                DImage(i - 1);
            }

            
            // Sauver le document
            nvDoc.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing,
                          ref missing, ref missing, ref missing, ref missing, ref missing,
                           ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing);

            // Fermer le document
             nvDoc.Close(ref missing, ref missing, ref missing);

            // Fermeture de word
             msWord.Quit(ref missing, ref missing, ref missing);
        }

        // Copie Le verbe que je transfère  dans le clipboard
        public void Verbe()
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Liste_des_verbes\ListeP5.txt");
            int randomLineNumber = r.Next(0, lines.Length - 1);
            string line = lines[randomLineNumber];
            
            DataObject clipData = new DataObject(DataFormats.Text, line);
            
            Clipboard.SetDataObject(clipData, false);
        }
        //Méthode qui rempli une case avec une verbe aléatoire
        public void VerbeTexte(int i)
        {
            Verbe();
            
            lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i]).Start;
            lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i]).End;
            nvDoc.Range(lstPosStartV[lstPosStartV.Count - 1], lstPosEndV[lstPosEndV.Count - 1]).Select();
            msWord.Selection.Paste();
            nvDoc.Bookmarks.Add((string)lstFieldsVerf[i], ref missing);
        }
        //Méthode qui rempli une case avec un pronom aléatoire
        public void PronomText(int i)
        {
            
            Pronom();

            lstPosStart[lstPosStart.Count - 1] = nvDoc.Bookmarks.get_Item(lstFields[i]).Start;
            lstPosEnd[lstPosEnd.Count - 1] = nvDoc.Bookmarks.get_Item(lstEndFields[i]).End;
            nvDoc.Range(lstPosStart[lstPosStart.Count - 1], lstPosEnd[lstPosEnd.Count - 1]).Select();
            msWord.Selection.Paste();
            nvDoc.Bookmarks.Add((string)lstEndFields[i], ref missing);
        }
        
        // Méthode qui rempli TOUTE la ligne avec des croix, pas finie
        public void DImage(int i)
        {
            Croix(); 

            lstPosStartI[lstPosStartI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgD[i]).Start;
            lstPosEndI[lstPosEndI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgF[i]).End;
            nvDoc.Range(lstPosStartI[lstPosStartI.Count - 1], lstPosEndI[lstPosEndI.Count - 1]).Select();
            msWord.Selection.Paste();
            nvDoc.Bookmarks.Add((string)lstFieldsImgF[i], ref missing);

        }

        //Choisi un pronom aléatoire et le copie dans le presse papier
        public void Pronom()
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Liste_des_verbes\Pronoms.txt");
            int randomLineNumber = r.Next(0, lines.Length - 1);
            string line = lines[randomLineNumber];

            DataObject clipData = new DataObject(DataFormats.Text, line);
            Clipboard.SetDataObject(clipData, false);
        }

        //Copie une croix noire dans le presse papier
        public void Croix()
        {
            Image img = Image.FromFile(@"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Image\croix.bmp");
            DataObject clipData = new DataObject(DataFormats.Bitmap,img);
            Clipboard.SetDataObject(clipData, false);
        }
        // Bouton qui quitte l'application.
        public void Verbe6()
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Liste_des_verbes\ListeP6.txt");
            int randomLineNumber = r.Next(0, lines.Length - 1);
            string line = lines[randomLineNumber];

            DataObject clipData = new DataObject(DataFormats.Text, line);

            Clipboard.SetDataObject(clipData, false);
        }

        public void VerbeTexte6(int i)
        {
            Verbe6();

            lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i]).Start;
            lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i]).End;
            nvDoc.Range(lstPosStartV[lstPosStartV.Count - 1], lstPosEndV[lstPosEndV.Count - 1]).Select();
            msWord.Selection.Paste();
            nvDoc.Bookmarks.Add((string)lstFieldsVerf[i], ref missing);
        }

        public void Verbe7()
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Liste_des_verbes\ListeP7.txt");
            int randomLineNumber = r.Next(0, lines.Length - 1);
            string line = lines[randomLineNumber];

            DataObject clipData = new DataObject(DataFormats.Text, line);

            Clipboard.SetDataObject(clipData, false);
        }

        public void VerbeTexte7(int i)
        {
            Verbe7();

            lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i]).Start;
            lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i]).End;
            nvDoc.Range(lstPosStartV[lstPosStartV.Count - 1], lstPosEndV[lstPosEndV.Count - 1]).Select();
            msWord.Selection.Paste();
            nvDoc.Bookmarks.Add((string)lstFieldsVerf[i], ref missing);
        }

        public void Verbe8()
        {
            string[] lines = System.IO.File.ReadAllLines(@"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Liste_des_verbes\ListeP8.txt");
            int randomLineNumber = r.Next(0, lines.Length - 1);
            string line = lines[randomLineNumber];

            DataObject clipData = new DataObject(DataFormats.Text, line);

            Clipboard.SetDataObject(clipData, false);
        }

        public void VerbeTexte8(int i)
        {
            Verbe8();

            lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i]).Start;
            lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i]).End;
            nvDoc.Range(lstPosStartV[lstPosStartV.Count - 1], lstPosEndV[lstPosEndV.Count - 1]).Select();
            msWord.Selection.Paste();
            nvDoc.Bookmarks.Add((string)lstFieldsVerf[i], ref missing);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            {

                // Connexion à Word
                msWord = new Microsoft.Office.Interop.Word.Application();
                msWord.Visible = false; // mettez cette variable à true si vous souhaitez visualiser les opérations.
                missing = System.Reflection.Missing.Value;

                // Attribuer le nom
                object fileName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Fiche_P6.docx";

                //Microsoft.Office.Interop.Word.Document nvDoc;

                // Tester s'il s'agit d'un nouveau document ou d'un document existant.
                if (System.IO.File.Exists((string)fileName))
                {
                    // ouvrir le document existant
                    nvDoc = msWord.Documents.Open(ref fileName, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing,
                                ref missing);
                }
                else
                {
                    // Choisir le template
                    object templateName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Template\Test.dotm";
                    // Créer le document
                    nvDoc = msWord.Documents.Add(ref templateName, ref missing, ref missing, ref missing);
                }

                lstFieldsVerf = new List<object> { };
                lstFields = new List<object> { };
                lstEndFields = new List<object> { };
                lstPosStart = new List<object> { };
                lstPosEnd = new List<object> { };
                lstPosStartV = new List<object> { };
                lstPosEndV = new List<object> { };
                lstFieldsVerd = new List<object> { };
                lstFieldsImgD = new List<object> { };
                lstFieldsImgF = new List<object> { };
                lstPosStartI = new List<object> { };
                lstPosEndI = new List<object> { };
                //boucle for qui rempli ligne par ligne le document 
                for (int i = 1; i <= 10; i++)
                {
                    lstFieldsVerd.Add("dverbe" + i);
                    lstFieldsVerf.Add("fverbe" + i);

                    lstEndFields.Add("fpronom" + i);
                    lstFields.Add("dpronom" + i);

                    lstFieldsImgD.Add("imaged" + i);
                    lstFieldsImgF.Add("imagef" + i);

                    lstPosStart.Add(new Object());
                    lstPosStart[lstPosStart.Count - 1] = nvDoc.Bookmarks.get_Item(lstFields[i - 1]).Start;
                    lstPosEnd.Add(new Object());
                    lstPosEnd[lstPosEnd.Count - 1] = nvDoc.Bookmarks.get_Item(lstEndFields[i - 1]).End;

                    lstPosStartV.Add(new Object());
                    lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i - 1]).Start;
                    lstPosEndV.Add(new Object());
                    lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i - 1]).End;

                    lstPosStartI.Add(new Object());
                    lstPosStartI[lstPosStartI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgD[i - 1]).Start;
                    lstPosEndI.Add(new Object());
                    lstPosEndI[lstPosEndI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgF[i - 1]).End;

                    PronomText(i - 1);
                    VerbeTexte6(i - 1);
                    DImage(i - 1);
                }


                // Sauver le document
                nvDoc.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing,
                              ref missing, ref missing, ref missing, ref missing, ref missing,
                               ref missing, ref missing, ref missing, ref missing, ref missing,
                                ref missing);

                // Fermer le document
                nvDoc.Close(ref missing, ref missing, ref missing);

                // Fermeture de word
                msWord.Quit(ref missing, ref missing, ref missing);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
                    // Connexion à Word
                    msWord = new Microsoft.Office.Interop.Word.Application();
                    msWord.Visible = false; // mettez cette variable à true si vous souhaitez visualiser les opérations.
                    missing = System.Reflection.Missing.Value;

                    // Attribuer le nom
                    object fileName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Fiche_P7.docx";

                    //Microsoft.Office.Interop.Word.Document nvDoc;

                    // Tester s'il s'agit d'un nouveau document ou d'un document existant.
                    if (System.IO.File.Exists((string)fileName))
                    {
                        // ouvrir le document existant
                        nvDoc = msWord.Documents.Open(ref fileName, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing, ref missing, ref missing,
                                    ref missing);
                    }
                    else
                    {
                        // Choisir le template
                        object templateName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Template\Test.dotm";
                        // Créer le document
                        nvDoc = msWord.Documents.Add(ref templateName, ref missing, ref missing, ref missing);
                    }

                    lstFieldsVerf = new List<object> { };
                    lstFields = new List<object> { };
                    lstEndFields = new List<object> { };
                    lstPosStart = new List<object> { };
                    lstPosEnd = new List<object> { };
                    lstPosStartV = new List<object> { };
                    lstPosEndV = new List<object> { };
                    lstFieldsVerd = new List<object> { };
                    lstFieldsImgD = new List<object> { };
                    lstFieldsImgF = new List<object> { };
                    lstPosStartI = new List<object> { };
                    lstPosEndI = new List<object> { };
                    //boucle for qui rempli ligne par ligne le document 
                    for (int i = 1; i <= 10; i++)
                    {
                        lstFieldsVerd.Add("dverbe" + i);
                        lstFieldsVerf.Add("fverbe" + i);

                        lstEndFields.Add("fpronom" + i);
                        lstFields.Add("dpronom" + i);

                        lstFieldsImgD.Add("imaged" + i);
                        lstFieldsImgF.Add("imagef" + i);

                        lstPosStart.Add(new Object());
                        lstPosStart[lstPosStart.Count - 1] = nvDoc.Bookmarks.get_Item(lstFields[i - 1]).Start;
                        lstPosEnd.Add(new Object());
                        lstPosEnd[lstPosEnd.Count - 1] = nvDoc.Bookmarks.get_Item(lstEndFields[i - 1]).End;

                        lstPosStartV.Add(new Object());
                        lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i - 1]).Start;
                        lstPosEndV.Add(new Object());
                        lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i - 1]).End;

                        lstPosStartI.Add(new Object());
                        lstPosStartI[lstPosStartI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgD[i - 1]).Start;
                        lstPosEndI.Add(new Object());
                        lstPosEndI[lstPosEndI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgF[i - 1]).End;

                        PronomText(i - 1);
                        VerbeTexte7(i - 1);
                        DImage(i - 1);
                    }


                    // Sauver le document
                    nvDoc.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing,
                                  ref missing, ref missing, ref missing, ref missing, ref missing,
                                   ref missing, ref missing, ref missing, ref missing, ref missing,
                                    ref missing);

                    // Fermer le document
                    nvDoc.Close(ref missing, ref missing, ref missing);

                    // Fermeture de word
                    msWord.Quit(ref missing, ref missing, ref missing);
                
            }

        private void button3_Click(object sender, EventArgs e)
        {
            // Connexion à Word
            msWord = new Microsoft.Office.Interop.Word.Application();
            msWord.Visible = false; // mettez cette variable à true si vous souhaitez visualiser les opérations.
            missing = System.Reflection.Missing.Value;

            // Attribuer le nom
            object fileName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Fiche_P8.docx";

            //Microsoft.Office.Interop.Word.Document nvDoc;

            // Tester s'il s'agit d'un nouveau document ou d'un document existant.
            if (System.IO.File.Exists((string)fileName))
            {
                // ouvrir le document existant
                nvDoc = msWord.Documents.Open(ref fileName, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing,
                            ref missing);
            }
            else
            {
                // Choisir le template
                object templateName = @"C:\Users\RochaixJe\Desktop\Atelier_Projet_Feuille_Exercice\Template\Test.dotm";
                // Créer le document
                nvDoc = msWord.Documents.Add(ref templateName, ref missing, ref missing, ref missing);
            }

            lstFieldsVerf = new List<object> { };
            lstFields = new List<object> { };
            lstEndFields = new List<object> { };
            lstPosStart = new List<object> { };
            lstPosEnd = new List<object> { };
            lstPosStartV = new List<object> { };
            lstPosEndV = new List<object> { };
            lstFieldsVerd = new List<object> { };
            lstFieldsImgD = new List<object> { };
            lstFieldsImgF = new List<object> { };
            lstPosStartI = new List<object> { };
            lstPosEndI = new List<object> { };
            //boucle for qui rempli ligne par ligne le document 
            for (int i = 1; i <= 10; i++)
            {
                lstFieldsVerd.Add("dverbe" + i);
                lstFieldsVerf.Add("fverbe" + i);

                lstEndFields.Add("fpronom" + i);
                lstFields.Add("dpronom" + i);

                lstFieldsImgD.Add("imaged" + i);
                lstFieldsImgF.Add("imagef" + i);

                lstPosStart.Add(new Object());
                lstPosStart[lstPosStart.Count - 1] = nvDoc.Bookmarks.get_Item(lstFields[i - 1]).Start;
                lstPosEnd.Add(new Object());
                lstPosEnd[lstPosEnd.Count - 1] = nvDoc.Bookmarks.get_Item(lstEndFields[i - 1]).End;

                lstPosStartV.Add(new Object());
                lstPosStartV[lstPosStartV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerd[i - 1]).Start;
                lstPosEndV.Add(new Object());
                lstPosEndV[lstPosEndV.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsVerf[i - 1]).End;

                lstPosStartI.Add(new Object());
                lstPosStartI[lstPosStartI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgD[i - 1]).Start;
                lstPosEndI.Add(new Object());
                lstPosEndI[lstPosEndI.Count - 1] = nvDoc.Bookmarks.get_Item(lstFieldsImgF[i - 1]).End;

                PronomText(i - 1);
                VerbeTexte8(i - 1);
                DImage(i - 1);
            }


            // Sauver le document
            nvDoc.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing,
                          ref missing, ref missing, ref missing, ref missing, ref missing,
                           ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing);

            // Fermer le document
            nvDoc.Close(ref missing, ref missing, ref missing);

            // Fermeture de word
            msWord.Quit(ref missing, ref missing, ref missing);
        }

        private void label1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
  }
  



