
// Version C#/XAML for HTML5 : CSharpHTML5

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace IAVB
{
    public sealed partial class MainPage : Page
    {
        // Mettre true tant qu'on n'a pas trouver comment faire lstBox.ScrollIntoView
        private const bool bAjoutListeDebut = true; // Sinon à la fin

        // Pour la démo sur les exemples
        private const int iNumLigneDebutExemples = 8;

        private const bool bVersionModifiee = true;
        private const bool bNormalisationSortieTrimEtVbLf = true;

        //public bool m_bVoixActive;

        //public bool m_bMSAgent;

        private IAVB.clsIAVB m_oIAVB = new IAVB.clsIAVB();

        public MainPage()
        {
            this.InitializeComponent();

            // Enter construction logic here...
            m_oIAVB.Initialiser(bVersionModifiee, bNormalisationSortieTrimEtVbLf);
            InitAssertionsExemples();
        }

        private void AjouterExemple(string sExemple)
        {
            this.ListAssert.Items.Add(sExemple);
        }

        private void InitAssertionsExemples()
        {
            AjouterExemple(m_oIAVB.sCmdEffBase + " ' Effacer toute la base");
            AjouterExemple(m_oIAVB.sCmdEff + " ' Effacer la denière assertion");
            AjouterExemple(m_oIAVB.sCmdEff + "1 ' Effacer l'assertion n°1");
            AjouterExemple(m_oIAVB.sCmdLister + " ' Lister les assertions de la base");
            AjouterExemple(m_oIAVB.sCmdCopier + " ' Copier la discussion dans le presse papier : commande non disponible dans la version C#2Html");
            //AjouterExemple(m_oIAVB.sCmdCopier + " ' Copier la discussion dans le presse papier");
            AjouterExemple(m_oIAVB.sCmdSilence + " ' Mettre en sourdine");
            AjouterExemple(m_oIAVB.sCmdParler + " ' Ré-activer la voix");

            List<string> lst = new List<string>();
            IAVB.clsInitIAVB ex = new IAVB.clsInitIAVB(lst);
            ex.InitAssertionsExemples();
            foreach (string sExemple in lst) AjouterExemple(sExemple);
        }

        private void CmdSuivant_click(object sender, RoutedEventArgs e)
        {
            if (ListAssert.SelectedIndex < iNumLigneDebutExemples)
                ListAssert.SelectedIndex = iNumLigneDebutExemples;
            else if (ListAssert.SelectedIndex < ListAssert.Items.Count - 1)
                ListAssert.SelectedIndex = ListAssert.SelectedIndex + 1;
            TraiterAssertion();
        }

        private void TraiterAssertion()
        {
            TextInput0.Text = (string)ListAssert.SelectedValue;
            TraiterCmd();
        }

        private void TraiterCmd()
        {
            bool bMemVoixActive = m_oIAVB.m_bVoixActive;

            // Si la zone de saisie est multiligne, faire un Trim
            string sTxt = TextInput0.Text.Trim();
            m_oIAVB.IAVBMain(sTxt);

            if (m_oIAVB.m_bCopierPressePapier)
            {
                if (bCopierPressePapier(m_oIAVB.m_sDiscussion))
                    m_oIAVB.CopierPressePapierOk();
                else
                    m_oIAVB.CopierPressePapierPb();
            }

            if (this.TextInput0.Text.Length == 0)
            {
                AjouterListIA("", bEnd:false);
                PositionnerListIA();
            }

            if (!string.IsNullOrEmpty(m_oIAVB.m_sRappelQuestion))
            {
                AjouterListIA(m_oIAVB.m_sRappelQuestion, bEnd: false);
                PositionnerListIA();
            }

            if (!string.IsNullOrEmpty(m_oIAVB.m_sReponse))
            {
                string[] asReponses = m_oIAVB.m_sReponse.Split(
                    new string[] { IAVB.clsIAVB.vbLf }, StringSplitOptions.None);
                foreach (string sReponse in asReponses)
                {
                    if (string.IsNullOrEmpty(sReponse.Trim())) continue;
                    AjouterListIA(sReponse, bEnd: false);
                }
                PositionnerListIA();
            }
        }

        public void AjouterListIA(string sTxt, bool bEnd = true)
        {
            if (bAjoutListeDebut)
            {
                ListBoxItem lbi = new ListBoxItem();
                lbi.Content = sTxt;
                lbi.HorizontalAlignment = Windows.UI.Xaml.HorizontalAlignment.Left;
                ListIA.Items.Insert(0, lbi);
            }
            else ListIA.Items.Add(sTxt);
        }

        public void PositionnerListIA()
        {
            if (bAjoutListeDebut)
                ListBoxSetSelectedIndex(ListIA, 0); // Au début
            else
                ListBoxSetSelectedIndex(ListIA, ListIA.Items.Count - 1); // A la fin
        }

        private void ListBoxSetSelectedIndex(ListBox lstBox, int index)
        {
            lstBox.SelectedIndex = index;
            lstBox.SelectedItem = lstBox.Items[lstBox.SelectedIndex];
            //lstBox.UpdateLayout();
            //lstBox.ScrollIntoView(lstBox.SelectedItem); // N'existe pas en C#HTML5
        }

        public static bool bCopierPressePapier(string sInfo)
        {
            return false; 
            // ToDo
            //try
            //{
            //    DataObject dataObj = new DataObject();
            //    dataObj.SetData(DataFormats.Text, sInfo);
            //    Clipboard.SetDataObject(dataObj);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    //ProjectData.SetProjectError(exception1);
            //    //Exception ex = exception1;
            //    //stringVB$t_string$S0 = "";
            //    //AfficherMsgErreur2(refex, "CopierPressePapier", "", "", false, r
            //    //ProjectData.ClearProjectError();
            //}
        }

        private void CmdGo_Click(object sender, RoutedEventArgs e)
        {
            TraiterCmd();
        }

        private void ListAssert_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TextInput0.Text = (string)ListAssert.SelectedValue;
        }

        //private void ListAssert_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        //{
        //    TraiterAssertion();
        //}

        private void ListIA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //throw new NotImplementedException();
        }

        // ToDo
        //private void TextInput0_KeyUp(object sender, KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter) TraiterCmd();
        //}

    }
}
