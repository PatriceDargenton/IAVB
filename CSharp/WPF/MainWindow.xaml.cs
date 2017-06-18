
using IAVB; // pour clsIAVB2.vbLf
//using Utils; // pour CopierPressePapier

using System;
using System.Collections.Generic;
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

using System.Speech;
using System.Speech.Synthesis;

//using System.Diagnostics; // Pour Debug.Write()

namespace IAVBWpf
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Pour la démo sur les exemples
        private const int iNumLigneDebutExemples = 8;

        private const bool bVersionModifiee = true;
        private const bool bNormalisationSortieTrimEtVbLf = true;

        public bool m_bVoixActive;
        
        private SpeechSynthesizer m_synth = new SpeechSynthesizer();

        private IAVB.clsIAVB m_oIAVB = new IAVB.clsIAVB();

        public MainWindow()
        {
            InitializeComponent();

            // En WPF, StartPosition n'existe pas, il faut passer par le xaml :
            //this.StartPosition = FormStartPosition.CenterScreen;
            //<Window x:Class="WpfApplication1.Window1" 
            //    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
            //    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
            //    Title="Window1" 
            //    Height="500" Width="500"
            //    WindowStartupLocation="Manual" 
            //    Left="0" Top="0">
            //</Window>

            Utils.clsAppUtils.AppTitle = clsIAVB.sIAVB;

            listParole.Items.Add(clsSynthVoc.sModeSilencieux);
            InitVoix();
            listParole.SelectedIndex = clsSynthVoc.iIndexModeSilencieux;

            m_oIAVB.Initialiser(bVersionModifiee, bNormalisationSortieTrimEtVbLf);
            InitAssertionsExemples();
        }

        private void InitVoix()
        {
            var lstParoles = new List<string>();
            clsSynthVoc.InitVoix(m_synth, lstParoles);
            foreach (string sParole in lstParoles) listParole.Items.Add(sParole);
        }

        private void AjouterExemple(string sExemple)
        {
            listAssert.Items.Add(sExemple);
        }

        private void InitAssertionsExemples()
        {
            AjouterExemple(m_oIAVB.sCmdEffBase + " ' Effacer toute la base");
            AjouterExemple(m_oIAVB.sCmdEff + " ' Effacer la denière assertion");
            AjouterExemple(m_oIAVB.sCmdEff + "1 ' Effacer l'assertion n°1");
            AjouterExemple(m_oIAVB.sCmdLister + " ' Lister les assertions de la base");
            AjouterExemple(m_oIAVB.sCmdCopier + " ' Copier la discussion dans le presse papier");
            AjouterExemple(m_oIAVB.sCmdSilence + " ' Mettre en sourdine");
            AjouterExemple(m_oIAVB.sCmdParler + " ' Ré-activer la voix");

            List<string> lst = new List<string>();
            IAVB.clsInitIAVB ex = new IAVB.clsInitIAVB(lst);
            ex.InitAssertionsExemples();
            foreach (string sExemple in lst) AjouterExemple(sExemple);
        }
        
        private void TraiterAssertion()
        {
            //TextInput.Text = listAssert.Text;
            textInput.Text = listAssert.SelectedValue.ToString();

            // Veiller à ce que la question soit bien affichée avant d'activer la synthèse vocale
            clsUtilWPF.DoEvents();

            TraiterCmd();
        }

        private void TraiterCmd()
        {
            bool bMemVoixActive = m_oIAVB.m_bVoixActive;
            
            // Eviter la réentrance dans les fonctions
            if (m_oIAVB.m_bVoixActive) Activation(bDesactiver: true);

            // Si la zone de saisie est multiligne, faire un Trim
            string sTxt = textInput.Text.Trim();
            m_oIAVB.IAVBMain(sTxt);

            if (m_oIAVB.m_bCopierPressePapier)
            {
                if (Utils.clsMessageUtils.bCopyToClipboard(m_oIAVB.m_sDiscussion))
                    m_oIAVB.CopierPressePapierOk();
                else
                    m_oIAVB.CopierPressePapierPb();
            }

            if (textInput.Text.Length == 0)
            {
                listIA.Items.Add("");
                PositionnerListIA();
            }

            // Activer ou réactiver la parole
            //if (!bMemVoixActive && m_oIAVB.m_bVoixActive) listParole.Text = sModeVocalDef;
            if (!bMemVoixActive && m_oIAVB.m_bVoixActive) listParole.SelectedValue = clsSynthVoc.sModeVocalDef;

            string sReponseVocale = m_oIAVB.m_sReponseVocale;
            if (sReponseVocale.Length > 0)
            {
                string[] asReponses = sReponseVocale.Split(
                    new string[] { clsIAVB.vbLf }, StringSplitOptions.None);
                int i = 0;
                foreach (string sReponse in asReponses)
                {
                    if (sReponse.Trim().Length == 0) continue;
                    if (i == 0 && asReponses.Length >= 1 && asReponses[1] == asReponses[0])
                    {
                        // Rappel de la question : dans ce cas on peut afficher le texte en 1er

                        // Réponse texte
                        if (sReponse.Trim().Length == 0) continue;
                        listIA.Items.Add(sReponse);
                        PositionnerListIA();

                        bDire(sReponse); // Réponse vocale
                    }
                    else if (i == 1 && asReponses.Length >= 1 && asReponses[1] == asReponses[0])
                    {
                        // Déjà traité
                    }
                    else if (i % 2 == 0) bDire(sReponse); // Réponse vocale
                    else
                    {
                        // Réponse texte
                        if (sReponse.Trim().Length == 0) continue;
                        listIA.Items.Add(sReponse);
                        PositionnerListIA();
                    }
                    i += 1;
                }

                // Désactiver la parole
                if (!m_oIAVB.m_bVoixActive) listParole.SelectedValue = clsSynthVoc.sModeSilencieux;

                goto Fin;
            }

            if (m_oIAVB.m_sRappelQuestion.Length > 0)
            {
                listIA.Items.Add(m_oIAVB.m_sRappelQuestion);
                PositionnerListIA();
            }

            if (m_oIAVB.m_sReponse.Length > 0)
            {
                string[] asReponses = m_oIAVB.m_sReponse.Split(
                    new string[] { clsIAVB.vbLf }, StringSplitOptions.None);
                foreach (string sReponse in asReponses)
                {
                    if (sReponse.Trim().Length == 0)
                        continue;
                    listIA.Items.Add(sReponse);
                }
                PositionnerListIA();
            }

        Fin:
            if (bMemVoixActive)
            {
                Activation();
                cmdSuivant.Focus();
            }

        }

        public bool bDire(string sParole)
        {
            //if (listParole.Text == sModeMSAgent) return bDireMSAgent(sParole);

            if (listParole.SelectedValue.ToString() == clsSynthVoc.sModeSilencieux) return false;

            if ((!(m_synth == null) && (clsSynthVoc.m_sVoix.Length > 0)))
            {
                while (m_synth.State == SynthesizerState.Speaking) Util.Attendre();

                if (m_synth.State == SynthesizerState.Ready)
                {
                    m_synth.SelectVoice(clsSynthVoc.m_sVoix);
                    Activation(bDesactiver: true);
                    // m_synth.SpeakAsync(sParole)
                    m_synth.Speak(sParole);
                    //  Attendre la fin
                    Activation();
                }

            }
            else listParole.SelectedValue = clsSynthVoc.sModeSilencieux;

            return true;
        }

        private void Activation(bool bDesactiver = false)
        {
            // ToDo
            //chkAuto.Enabled = true;
            //cmdSuivant.Enabled = true;
            //listAssert.Enabled = !bDesactiver;
            //listIA.Enabled = !bDesactiver;
            //textInput.Enabled = !bDesactiver;
            //cmdGo.Enabled = !bDesactiver;

            // Ici DoEvents permet d'afficher la répétition de la question avant la réponse d'IAVB
            // En Winform ça marche bien, par contre en WPF avec la synthèse vocale
            //  ça marche moins bien : mieux vaut ne pas faire de DoEvents, pas grave 
            //  (on peut quand même voir la question dans la zone des commandes)
            // WinForm : Application.DoEvents(); // Requis en C#
            //clsUtilWPF.DoEvents(); // Requis en C#, pas en VB
        }

        public void PositionnerListIA()
        {
            // Toujours se positionner sur la dernière ligne
            //listIA.SelectedIndex = listIA.Items.Count - 1;
            WPFListBoxSetSelectedIndex(listIA, listIA.Items.Count - 1);
        }

        private static void WPFListBoxSetSelectedIndex(ListBox lstBox, int index) 
        {
            lstBox.SelectedIndex = index;
            // How do you programmatically set focus to the SelectedItem in a WPF ListBox that already has focus?
            // http://stackoverflow.com/questions/10444518/how-do-you-programmatically-set-focus-to-the-selecteditem-in-a-wpf-listbox-that
            lstBox.SelectedItem = lstBox.Items[lstBox.SelectedIndex];
            lstBox.UpdateLayout();
            lstBox.ScrollIntoView(lstBox.SelectedItem);

            // Pb : il faudrait restaurer le focus après coup
            //  car par exemple, on veut conserver le focus sur le bouton CmdSuivant
            //var listBoxItem = (ListBoxItem)lstBox
            //    .ItemContainerGenerator.ContainerFromItem(lstBox.SelectedItem);
            //if (listBoxItem != null) listBoxItem.Focus();
            //lstBox.Focus();
        }

        #region "Gestion des événements"
        
        private void cmdGo_Click(object sender, RoutedEventArgs e)
        {
            TraiterCmd();
        }
        
        private void cmdSuivant_click(object sender, RoutedEventArgs e)
        {
            if (listAssert.SelectedIndex < iNumLigneDebutExemples)
            {
                //listAssert.SelectedIndex = iNumLigneDebutExemples;
                WPFListBoxSetSelectedIndex(listAssert, iNumLigneDebutExemples);
            }
            else if (listAssert.SelectedIndex < listAssert.Items.Count - 1)
            {
                //listAssert.SelectedIndex = listAssert.SelectedIndex + 1;
                WPFListBoxSetSelectedIndex(listAssert, listAssert.SelectedIndex + 1);
            }
            TraiterAssertion();
        }

        private void listAssert_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //TextInput.Text = listAssert.Text;
            textInput.Text = listAssert.SelectedValue.ToString();
        }

        private void listAssert_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TraiterAssertion();
        }

        private void listIA_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //textInput.Text = listIA.Text;
            textInput.Text = listIA.SelectedValue.ToString();
        }

        // Pas au point : éviter de traiter l'événement SelectionChanged après le DoubleClick !
        //private void listIA_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        //{
        //    //textInput.Text = listIA.Text;
        //    textInput.Text = listIA.SelectedValue.ToString();
        //    TraiterCmd();
        //}

        private void textInput_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) TraiterCmd();
        }

        private void listParole_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (listParole.SelectedValue.ToString() == clsSynthVoc.sModeSilencieux) m_oIAVB.m_bVoixActive = false;
            else m_oIAVB.m_bVoixActive = true;
            clsSynthVoc.LireVoix(listParole.SelectedIndex, listParole.SelectedValue.ToString());
        }

        #endregion
    }
    
}
