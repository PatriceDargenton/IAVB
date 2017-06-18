
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Speech;
using System.Speech.Synthesis;

namespace IAVB
{
    class clsSynthVoc
    {
        public const string sVoixFR = "Microsoft Hortense Desktop";
        public const string sVoix_enUSWin10 = "Microsoft Zira Desktop";
        public const string sVoix_enUSWin7 = "Microsoft Anna";
        public const string sFrancais = "Français (Hortense)";
        public const string sAnglais = "Anglais (Zira)";
        public const string sModeVocalDef = sFrancais;
        public const string sModeSilencieux = "Silencieux";
        //public const string sModeMSAgent = "MS-Agent";
        public const int iIndexModeSilencieux = 0;
        public static string m_sVoix = "";

        public static void InitVoix(SpeechSynthesizer m_synth, List<string> m_lstParoles)
        {
            //  SpeechSynthesizer.GetInstalledVoices, methode (CultureInfo)
            //  https://msdn.microsoft.com/fr-fr/library/ms586870(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
            //  Output information about all of the installed voices that
            //   support the en-US locacale. 
            Dictionary<string, VoiceInfo> dicoVoix = new Dictionary<string, VoiceInfo>();
            foreach (InstalledVoice voix in m_synth.GetInstalledVoices()) // New CultureInfo("en-US"))
            {
                VoiceInfo info = voix.VoiceInfo;
                string sCle = info.Name + " (" + info.Culture.Name + ", " +
                    info.Gender.ToString() + ")";
                if (info.Name == sVoixFR) sCle = sFrancais;
                if (info.Name == sVoix_enUSWin10) sCle = sAnglais;
                if (info.Name == sVoix_enUSWin7) sCle = sAnglais;
                if (!dicoVoix.ContainsKey(sCle))
                {
                    dicoVoix.Add(sCle, info);
                    m_lstParoles.Add(sCle);
                }
            }
        }

        public static void LireVoix(int iIndexVoix, string sLangue)
        {
            if (iIndexVoix < iIndexModeSilencieux) return;
            
            string sVoix = "";
            if (sLangue == clsSynthVoc.sFrancais) sVoix = clsSynthVoc.sVoixFR;
            else if (sLangue == clsSynthVoc.sAnglais) sVoix = clsSynthVoc.sVoix_enUSWin10;
            else
            {
                sVoix = sLangue;
                int iPos = sVoix.IndexOf("(");
                if (iPos > -1) sVoix = sVoix.Substring(0, (iPos - 1)).Trim();
            }
            m_sVoix = sVoix;
        }

    }
}
