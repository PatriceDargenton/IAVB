
using IAVB; // pour clsIAVB2.vbLf
//using LibIAVB2.clsIAVB2; // pour vbLf : Ne marche pas

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Speech;
using System.Speech.Synthesis;

//using System.Diagnostics; // Pour Debug.Write()

namespace IAVB
{
    public partial class frmIAVB : Form
    {
        private bool bDebug = IAVB.Program.bDebug;

        // Pour la démo sur les exemples
        private const int iNumLigneDebutExemples = 8;

        private const bool bVersionModifiee = true;
        private const bool bNormalisationSortieTrimEtVbLf = true;

        public bool m_bVoixActive;

        private SpeechSynthesizer m_synth = new SpeechSynthesizer();

        private IAVB.clsIAVB m_oIAVB = new IAVB.clsIAVB();

        #region "Initialisations"

        public frmIAVB()
        {
            InitializeComponent();
            if (bDebug) this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void frmIAVB_Load(object sender, EventArgs e)
        {
            Initialisations();
        }

        private void Initialisations()
        {
            string sTxt = Text + " - V" + IAVB.Program.sVersionAppli + 
                " (" + IAVB.Program.sDateVersionAppli + ")";
            if (bDebug) sTxt += " - Debug";
            this.Text = sTxt;

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
            foreach(string sParole in lstParoles)
            {
                listParole.Items.Add(sParole);
            }
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
            foreach (string sExemple in lst)
            {
                AjouterExemple(sExemple);
            }

        }

        #endregion
        
        private void TraiterAssertion()
        {
            textInput.Text = listAssert.Text;
            
            // Veiller à ce que la question soit bien affichée avant d'activer la synthèse vocale
            Application.DoEvents();

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
            if (!bMemVoixActive && m_oIAVB.m_bVoixActive) listParole.Text = clsSynthVoc.sModeVocalDef;

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
                if (!m_oIAVB.m_bVoixActive) listParole.Text = clsSynthVoc.sModeSilencieux;

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

            if (listParole.Text == clsSynthVoc.sModeSilencieux) return false;

            if ((!(m_synth == null) && (clsSynthVoc.m_sVoix.Length > 0))) 
            {
                while (m_synth.State == SynthesizerState.Speaking) Util.Attendre();
            
                if (m_synth.State == SynthesizerState.Ready) {
                    m_synth.SelectVoice(clsSynthVoc.m_sVoix);
                    Activation(bDesactiver:true);
                    // m_synth.SpeakAsync(sParole)
                    m_synth.Speak(sParole);
                    //  Attendre la fin
                    Activation();
                }
            
            }
            else listParole.Text = clsSynthVoc.sModeSilencieux;
        
            return true;
        }
        
        private void Activation(bool bDesactiver= false) 
        {
            //chkAuto.Enabled = true;
            cmdSuivant.Enabled = true;
            listAssert.Enabled = !bDesactiver;
            listIA.Enabled = !bDesactiver;
            textInput.Enabled = !bDesactiver;
            cmdGo.Enabled = !bDesactiver;

            // Ici DoEvents permet d'afficher la répétition de la question avant la réponse d'IAVB
            Application.DoEvents(); // Requis en C#, pas en VB
        }

        public void PositionnerListIA()
        {
            // Toujours se positionner sur la dernière ligne
            listIA.SelectedIndex = listIA.Items.Count - 1;
        }

        #region "Gestion des événements"

        private void cmdGo_Click(object sender, EventArgs e)
        {
            TraiterCmd();
        }

        private void cmdSuivant_Click(object eventSender, EventArgs eventArgs)
        {
            if (listAssert.SelectedIndex < iNumLigneDebutExemples)
                listAssert.SelectedIndex = iNumLigneDebutExemples;
            else if (listAssert.SelectedIndex < listAssert.Items.Count - 1)
                listAssert.SelectedIndex = listAssert.SelectedIndex + 1;
            TraiterAssertion();
        }

        private void listAssert_SelectedIndexChanged(object sender, EventArgs e)
        {
            textInput.Text = listAssert.Text;
        }

        private void listAssert_DoubleClick(object sender, EventArgs e)
        {
            TraiterAssertion();
        }
        
        private void listIA_SelectedIndexChanged(object sender, EventArgs e)
        {
            textInput.Text = listIA.Text;
        }

        private void listIA_DoubleClick(object sender, EventArgs e)
        {
            textInput.Text = listIA.Text;
            TraiterCmd();
        }

        void textInput_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Enter)) TraiterCmd();
        }

        private void listParole_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listParole.Text == clsSynthVoc.sModeSilencieux) m_oIAVB.m_bVoixActive = false;
            else m_oIAVB.m_bVoixActive = true;
            clsSynthVoc.LireVoix(listParole.SelectedIndex, listParole.Text);
        }

        #endregion

    }
}
