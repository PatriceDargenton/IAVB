
// Fichier clsIAVB.cs : Classe du moteur IAVB : Intelligence Artificielle Vraiment Basique
// ------------------

// Documentation : IAVB.html
// http://patrice.dargenton.free.fr/ia/iavb/IAVB.html
// https://github.com/PatriceDargenton/IAVB
// Version 3.12 du 25/05/2017

// IAVB est la transcription en Visual Basic d'un logiciel paru dans la revue
//  MICRO-SYSTEMES en Décembre 1984, pages 195-202 :
//  "Mini-système expert pour Apple II" par Philippe LARVET :
//  Gestion d'une base de connaissances.

// Par Patrice Dargenton : mailto:patrice.dargenton@free.fr
// http://patrice.dargenton.free.fr/index.html
// http://patrice.dargenton.free.fr/CodesSources/index.html

// Conventions de nommage des variables :
// ------------------------------------
// b pour Boolean (booléen vrai ou faux)
// i pour Integer
// l pour Long
// r pour nombre Réel (Single, Double ou Decimal)
// s pour String
// c pour Char ou Byte
// d pour Date
// u pour Unsigned (non signé : entier positif)
// a pour Array (tableau) : ()
// m_ pour variable Membre de la classe ou de la feuille (Form)
//  (mais pas pour les constantes)
// frm pour Form
// cls pour Classe
// mod pour Module
// ...
// ------------------------------------

using System;
using System.Collections.Generic;
using System.Text;

namespace IAVB
{
public class clsIAVB
{
#region Interface

    // Booléen pour forcer le fonctionnement dans la version originale (1984)
    //  ou bien modifiée (2001)
    public bool m_bVersionModifiee = true;

    // Normaliser les espaces et sauts de ligne en sortie
    public bool m_bNormalisationSortieTrimEtVbLf = true;

    // Convertir tout en minuscules (pour éviter d'avoir à taper des majuscules avec accent,
    //  et pour pouvoir prononcer "ça" via MS-Agent)
    public const bool bTraiterEnMinuscules = true;

    // iNbMotsBCMax : Nbre de mots signifiants susceptibles
    //  d'être conservé en mémoire pour chaque assertion (au moins 3).
    //  seul les 2 premiers mots peuvent former une clé pour une relation H ou une indirection
    private const int iNbMotsBCMax = 4;

    // Nbre de mots signifiants de l'entrée max.
    // Peut être > iNbMotsBCMax notamment pour les compositions de fonctions
    private const int iNbMotsEntreeMax = 10;

#endregion

#region Constantes

    public const string sIAVB = "IAVB 3.12";
    public const string vbLf = "\n";

    private const string sCmdCopieOk = "La discusion a été copiée dans le presse papier.";
    private const string sCmdCopiePb = "La discusion n'a pas pu être copiée dans le presse papier !";
    private const string sCmdEffacerConfirm = "Voulez-vous effacer toute la base ?";
    private const string sCmdEffacerOk = "Base entièrement effacée !";

    private const string sCmdParlerOk = "Je vous écoute.";
    private const string sCmdSilenceOk = "Okay, je la ferme.";

    private const string sCmdSupprAssertionOk1 = "Assertion ";
    private const string sCmdSupprAssertionOk2 = " supprimée.";
    private const string sCmdSupprAssertionPb = "Numéro de l'assertion hors limites !";
    
    private const string sErrAssertionDejaConnue = "Assertion déjà connue !";
    private const string sErrBaseVide = "Base vide !";
    private const string sErrCmdInconnue = "Commande inconnue !";
    private const string sErrPhraseIncomplete = "Votre phrase est incomplète !";
    private const string sErrPhraseTropCourte = "Votre phrase est trop courte !";
    private const string sErrPhraseTropLongue = "Votre phrase est trop longue !";
    private const string sReponseCompris = "Compris.";
    private const string sReponseConnaisPas = "Je ne connais pas ";
    private const string sReponseIgnore = "Je l'ignore.";
    private const string sReponseNon = "Non.";
    private const string sReponseOui = "Oui.";
    private const string sReponseRechApprof = "Recherche approfondie...";
    private const string sReponseRienConclure = "Je ne peux rien conclure !";

    private const string sDeuxPoints = ":";
    private const string sEgal = "=";
    private const string sEspace = " ";
    private const string sGM = "\"";
    private const string sParenthFerm = ")";
    private const string sParenthOuv = "(";
    private const string sPointExclam = "!";
    private const string sPointInterrog = "?";
    private const string sQuote = "'";
    private const string sSlash = "/";
    private const string sTiret = "-";

    private const string sTitreMsg = "Moteur IAVB3";

#endregion

#region Déclarations
        
    private List<TAssertion> m_lstBC; // Base de Connaissances (KB)

    // Liste des mots ignorés en général (non-signifiants)
    private List<string> m_lstMotsIgnores;

    public bool m_bCopierPressePapier;
    private bool m_bQuestionPreced_Donc;
    public bool m_bVoixActive;
    private int m_iNumAssertionEnCours;
    
    // Préparer les réponses multiples pour les renvoyer en une seule chaine de caractères
    private List<string> m_lstReponses;
    // Pour pouvoir copier la discussion dans le presse papier
    private List<string> m_lstReponsesTot;
    // Pour pouvoir synchroniser les réponses vocales avec les réponses écrites
    private List<string> m_lstReponsesVocales;

    public string m_sDiscussion, m_sRappelQuestion, m_sReponse, m_sReponseVocale;
    public string sCmdCopier, sCmdEffBase, sCmdEff, sCmdLister, sCmdParler, sCmdSilence;
    private string sMotDonc, sMotEst, sMotEstComme, sMotOr, sMotQU, sMotQuelque, sMotQuelques,
        sMotSont, sMotTou, sMotUn;

#endregion

#region Classes

    private class TAssertion
    {
        public int iNbMots;
        public string[] asMot;
        public int iPosFinMot1;
        public int iPosFinMot2;
        public string sEntree;
    }

    private class TParamEntree
    {
        public bool bMotQuestion;
        public bool bInterrogation;
        public bool bFinTraitement;
        public string sEntree;
        public string sEntreeCompilee; // Concaténation de tous les asMotsEntree(i)
        public int iNbMotsEntree; // Nbre de mots signifiants extraits de sEntree
        public string[] asMotsEntree;
        public int iPosFinMot1;
        public int iPosFinMot2;
        
        // n° de l'assertion contenant le mot signifiant n°K de l'entrée,
        //  tableau des K mots de l'entrée
        //  D'abord la première assertion : ControleExistenceMots(),
        //  puis les assertions suivantes : bAssertCMot() pour les indirections.
        public int[] aiNumAssertCMot;

        // Conservation de l'indice minimum pour
        //  l'espace de recherche max. de l'assertion
        //  (contenant un des mots de l'assertion en cours)
        public int iNumAssertMinRech;
    }
    
#endregion

#region Initialisations

    public clsIAVB()
    {
        m_lstReponsesTot = new List<string>();
        m_sRappelQuestion = "";
        m_sReponse = "";
        m_bCopierPressePapier = false;
        m_sDiscussion = "";
        m_sReponseVocale = "";
    }

    public void Initialiser(bool bVersionModifiee, bool bNormalisationSortieTrimEtVbLf)
    {

        m_bVersionModifiee = bVersionModifiee;
        m_bNormalisationSortieTrimEtVbLf = bNormalisationSortieTrimEtVbLf;

        m_lstMotsIgnores = new List<string>();

        m_lstBC = new List<TAssertion>();

        InitCmd();
        InitMotsIgnores();
    }

    private void InitAssertion(int iNumAssert)
    {
        m_lstBC[iNumAssert].asMot = new string[iNbMotsBCMax+1];
    }

    private void InitCmd()
    {
        sCmdEffBase = "/EFF";  // Effacer toute la base
        sCmdEff = "/D";        // Effacer la denière assertion, ou /Dn : l'assertion n° n"
        sCmdLister = "/L";     // Lister les assertions de la base
        sCmdCopier = "/C";     // Copier la discussion dans le presse papier
        sCmdSilence = "/S";    // Mettre en sourdine
        sCmdParler = "/PARLE"; // Ré-activer la voix
        sMotQU = "QU";
        sMotDonc = "DONC";
        sMotEstComme = "EST COMME";
        sMotEst = "EST";
        sMotSont = "SONT";
        sMotTou = "TOU"; // (TOUT, TOUS, TOUTES)
        sMotUn = "UN";
        sMotOr = "OR";
        sMotQuelque = "QUELQUE";
        sMotQuelques = "QUELQUES";
        if (bTraiterEnMinuscules)
        {
            sCmdEffBase = sCmdEffBase.ToLowerInvariant();
            sCmdEff = sCmdEff.ToLowerInvariant();
            sCmdLister = sCmdLister.ToLowerInvariant();
            sCmdCopier = sCmdCopier.ToLowerInvariant();
            sCmdSilence = sCmdSilence.ToLowerInvariant();
            sCmdParler = sCmdParler.ToLowerInvariant();
            sMotQU = sMotQU.ToLowerInvariant();
            sMotDonc = sMotDonc.ToLowerInvariant();
            sMotEstComme = sMotEstComme.ToLowerInvariant();
            sMotEst = sMotEst.ToLowerInvariant();
            sMotSont = sMotSont.ToLowerInvariant();
            sMotTou = sMotTou.ToLowerInvariant();
            sMotUn = sMotUn.ToLowerInvariant();
            sMotOr = sMotOr.ToLowerInvariant();
            sMotQuelque = sMotQuelque.ToLowerInvariant();
            sMotQuelques = sMotQuelques.ToLowerInvariant();
        }
    }

    private void InitMotsIgnores()
    {
        // Mots ignorés (car non-signifiants)
        m_lstMotsIgnores = new List<string> {
            "EST", "LE", "LA", "DE", "UN", "UNE", "OR", "L", "DU", "D",
            "LES", "DES", "ET", "QU", "QUE", "QUI", "SONT", "IL", "ELLE",
            "A", "T", "ÉTÉ", // "ETE"
            "EN", "OU", "COMMENT", "AU", "N", "NE", "S", "SE", "ÉTAIT", // "ETAIT"
            "QUOI", "C", "CE", "QUEL", "QUELLE", "QUELS", "QUELLES",
            "PAR", "LEQUEL", "LAQUELLE", "ÇA", "SIGNIFIE",
            "TOUT", "TOUS", "TOUTE", "TOUTES", "(", ")", "=", "Y"
        };

        // Convertir tout en minuscules
        if (bTraiterEnMinuscules)
            for (int i = 0; i < m_lstMotsIgnores.Count; i++)
                m_lstMotsIgnores[i] = m_lstMotsIgnores[i].ToLowerInvariant();
    }

#endregion

#region Moteur IAVB
    
    public void IAVBMain(string sEntree)
    {
        m_sReponseVocale = "";
        m_sDiscussion = "";
        m_bCopierPressePapier = false;
        m_sRappelQuestion = "";
        m_sReponse = "";
        m_lstReponses = new List<string>();
        m_lstReponsesVocales = new List<string>();
        var prm = new TParamEntree();
        prm.sEntree = sEntree;

        AnalyseEntree(prm);

        if (prm.bFinTraitement) return;

        if (prm.bInterrogation && prm.asMotsEntree[1] == sMotDonc)
        {
            bool bRechecheApprofondie = false;
            if (m_bVersionModifiee && m_bQuestionPreced_Donc)
            {
                bRechecheApprofondie = true;
                if (m_iNumAssertionEnCours > 0)
                {
                    String s1 = m_lstBC[m_iNumAssertionEnCours].sEntree + " :";
                    ReponseIAVB(s1, bListe:true);
                }
                ReponseIAVB(sReponseRechApprof, bAffiche2Etoiles: true, bListe: true);
            }
            Syllogisme(bRechecheApprofondie);
            m_bQuestionPreced_Donc = true;
            return;
        }
        
        m_bQuestionPreced_Donc = false;
        if (!prm.bInterrogation)
        {
            if (prm.iNbMotsEntree <= 1)
            {
                ReponseIAVB(sErrPhraseTropCourte, bAffiche2Etoiles: true);
                return;
            }
            AjoutBase(prm);
            return;
        }

        // Interrogation : controle existence des mots
        ControleExistenceMots(prm);

        if (prm.bFinTraitement) return;

        // Composition de fcts (relation verticale)
        if (prm.iNbMotsEntree <= iNbMotsBCMax)
            if (bRelationHorizontale(prm)) return;

        // Composition de fcts (relation verticale)
        if (bComposFonction(prm)) return;

        if (bIndirection(prm)) return;

        // Echec Final
        if (prm.bMotQuestion)
        {
            ReponseIAVB(sReponseIgnore, bAffiche2Etoiles: true);
            return;
        }

        ReponseIAVB(sReponseNon, bAffiche2Etoiles: true);
    }

    private void AjoutBase(TParamEntree prm)
    {
        // Assertion : Controle existence de l'assertion dans la base
        int m_iNbAssertions = m_lstBC.Count;
        for (int i = 0; i < m_iNbAssertions; i++)
        {
            string sGdTerme = "";
            int j = 1;
            do
            {
                sGdTerme = sGdTerme + m_lstBC[i].asMot[j];
                j++;
            }
            while (j <= iNbMotsBCMax);
            if (sGdTerme == prm.sEntreeCompilee)
            {
                if (m_bVersionModifiee) m_iNumAssertionEnCours = i;
                ReponseIAVB(sErrAssertionDejaConnue, bAffiche2Etoiles: true);
                return;
            }
        }

        AjouterAssertion();
        int i2 = m_lstBC.Count - 1;
        m_lstBC[i2].asMot = new string[iNbMotsBCMax + 1];

        int j2 = 1;
        do
        {
            m_lstBC[i2].asMot[j2] = prm.asMotsEntree[j2];
            j2++;
        }
        while (j2 <= iNbMotsBCMax);
        m_lstBC[i2].sEntree = prm.sEntree;
        m_lstBC[i2].iNbMots = prm.iNbMotsEntree;
        m_lstBC[i2].iPosFinMot1 = prm.iPosFinMot1;
        m_lstBC[i2].iPosFinMot2 = prm.iPosFinMot2;
        ReponseIAVB(sReponseCompris, bAffiche2Etoiles: true);
        m_iNumAssertionEnCours = m_lstBC.Count - 1;
    }

    private void AjouterAssertion()
    {
        var assert = new TAssertion();
        assert.asMot = new string[iNbMotsBCMax+1];
        for (int i = 0; i <= iNbMotsBCMax; i++) assert.asMot[i] = "";
        m_lstBC.Add(assert);
    }

    private void AnalyseEntree(TParamEntree prm)
    {
        prm.bInterrogation = false;
        prm.bInterrogation = false;
        prm.bFinTraitement = false;
        prm.bMotQuestion = false;
        prm.iNbMotsEntree = 0;
        prm.sEntreeCompilee = "";
        prm.asMotsEntree = new string[11];
        prm.aiNumAssertCMot = new int[11];

        for (int k = 1; k < iNbMotsBCMax +1 ; k++)
        {
            prm.asMotsEntree[k] = "";
            prm.aiNumAssertCMot[k] = 0;
        }

        int iLenEntree = clsVBUtil.iVBLen(prm.sEntree);
        if (iLenEntree == 0)
        {
            ReponseIAVB(prm.sEntree, bSilence: true);
            prm.bFinTraitement = true;
            return;
        }

        if (clsVBUtil.sVBLeft(prm.sEntree, 1) == sQuote)
        {
            ReponseIAVB(prm.sEntree, bSilence: true);
            prm.bFinTraitement = true;
            return;
        }

        if (bTraiterEnMinuscules && prm.sEntree != null) prm.sEntree = prm.sEntree.ToLower();

        if (clsVBUtil.sVBLeft(prm.sEntree, 1) == sSlash)
        {
            ReponseIAVB(prm.sEntree, bSilence: true);
            TraiterCmd(prm, iLenEntree);
            return ;
        }

        ReponseIAVB(prm.sEntree, bRappelReponse: true);
        if (clsVBUtil.sVBRight(prm.sEntree, 1) == sPointInterrog)
        {
            prm.sEntree = clsVBUtil.sVBLeft(prm.sEntree, iLenEntree - 1);
            iLenEntree = clsVBUtil.iVBLen(prm.sEntree);
            prm.bInterrogation = true;
        }

        AnalyseMots(prm, iLenEntree);
    }

    private void AnalyseMots(TParamEntree prm, int iLenEntree)
    {
        int iPos = 1;
        do
        {
            int iMemPos = iPos;
            string sLettreEntree = "";
            do
            {
                iPos++;
                sLettreEntree = clsVBUtil.sVBMid(prm.sEntree, iPos, 1);
            }
            while (iPos <= iLenEntree &&
                   sLettreEntree != sEspace &&
                   sLettreEntree != sQuote &&
                   sLettreEntree != sTiret &&
                   sLettreEntree != sParenthOuv &&
                   sLettreEntree != sParenthFerm);

            string sMotEntree = clsVBUtil.sVBMid(prm.sEntree, iMemPos, iPos - iMemPos);
            if (iMemPos == 1 && clsVBUtil.sVBLeft(sMotEntree, 2) == sMotQU)
                prm.bMotQuestion = true;
            else
            {
                bool bMotIgnore = false;
                for (int k = 0; k < m_lstMotsIgnores.Count; k++)
                {
                    if (m_lstMotsIgnores[k] == sMotEntree)
                    {
                        bMotIgnore = true;
                        break;
                    }
                }

                if (!bMotIgnore)
                {
                    // Ajout d'un mot signifiant
                    prm.iNbMotsEntree++;
                    // Modification : test aussi avec iNbMotsEntreeMax
                    if ((!prm.bInterrogation && prm.iNbMotsEntree > iNbMotsBCMax) ||
                        prm.iNbMotsEntree > iNbMotsEntreeMax)
                    {
                        ReponseIAVB(sErrPhraseTropLongue, bAffiche2Etoiles: true);
                        prm.bFinTraitement = true;
                        return;
                    }
                    prm.sEntreeCompilee = prm.sEntreeCompilee + sMotEntree;
                    prm.asMotsEntree[prm.iNbMotsEntree] = sMotEntree;
                    if (prm.iNbMotsEntree == 1) prm.iPosFinMot1 = iPos;
                    if (prm.iNbMotsEntree == 2) prm.iPosFinMot2 = iPos;
                }
            }
            if (iPos <= iLenEntree)
            {
                do iPos++;
                while (clsVBUtil.sVBMid(prm.sEntree, iPos, 1) == sEspace && iPos <= iLenEntree);
            }
        } while (iPos <= iLenEntree);

        if (string.IsNullOrEmpty(prm.asMotsEntree[1]))
        {
            ReponseIAVB(sErrPhraseIncomplete, bAffiche2Etoiles: true);
            prm.bFinTraitement = true;
        }

    }

    private void TraiterCmd(TParamEntree prm, int iLenEntree)
    {
        if (clsVBUtil.sVBLeft(prm.sEntree, sCmdLister.Length) == sCmdLister)
        { ListerBase(prm); return; }
        if (clsVBUtil.sVBLeft(prm.sEntree, sCmdEffBase.Length) == sCmdEffBase)
        { EffacerBase(prm); return; }
        if (clsVBUtil.sVBLeft(prm.sEntree, sCmdEff.Length) == sCmdEff)
        { SupprimerAssertion(prm, iLenEntree); return; }
        if (clsVBUtil.sVBLeft(prm.sEntree, sCmdCopier.Length) == sCmdCopier)
        { CopierPressePapier(prm); return; }

        if (clsVBUtil.sVBLeft(prm.sEntree, sCmdSilence.Length) == sCmdSilence)
        {
            ReponseIAVB(sCmdSilenceOk, bAffiche2Etoiles: true);
            m_bVoixActive = false;
            prm.bFinTraitement = true;
            return;
        }

        if (clsVBUtil.sVBLeft(prm.sEntree, sCmdParler.Length) == sCmdParler)
        {
            m_bVoixActive = true;
            ReponseIAVB(sCmdParlerOk, bAffiche2Etoiles: true);
            prm.bFinTraitement = true;
            return;
        }

        ReponseIAVB(sErrCmdInconnue, bAffiche2Etoiles: true);
        prm.bFinTraitement = true;
    }

    private void ListerBase(TParamEntree prm)
    {
        int iNbAssertions = m_lstBC.Count;
        if (iNbAssertions == 0)
        {
            ReponseIAVB(sErrBaseVide, bAffiche2Etoiles: true);
            prm.bFinTraitement = true;
            return;
        }
        for (int i = 0; i < iNbAssertions; i++)
        {
            String s1 = (i+1).ToString() + sEspace + m_lstBC[i].sEntree;
            ReponseIAVB(s1, bListe:true);
        }
        ReponseIAVB("", bFinListe:true);
        prm.bFinTraitement = true;
    }

    private void CopierPressePapier(TParamEntree prm)
    {
        StringBuilder sbDisc = new StringBuilder();
        foreach (string sReponse in m_lstReponsesTot) sbDisc.AppendLine(sReponse);
        m_sDiscussion = sbDisc.ToString();
        m_bCopierPressePapier = true;
        //ReponseIAVB(sCmdCopieOk, bAffiche2Etoiles: true);
        prm.bFinTraitement = true;
        return;
    }

    public void CopierPressePapierOk()
    {
        ReponseIAVB(sCmdCopieOk, bAffiche2Etoiles: true);
    }

    public void CopierPressePapierPb()
    {
        ReponseIAVB(sCmdCopiePb, bAffiche2Etoiles: true);
    }

    private void SupprimerAssertion(TParamEntree prm, int iLenEntree)
    {
        int iNbAssertions = m_lstBC.Count;
        int i = iNbAssertions;
        if (i == 0)
        {
            ReponseIAVB(sErrBaseVide, bAffiche2Etoiles: true);
            prm.bFinTraitement = true;
            return;
        }
        i -= 1; // Par défaut effacer la dernière assertion
        
        if (clsVBUtil.sVBMid(prm.sEntree, 3, 1) != sEspace)
        {
            string sNumAssert = clsVBUtil.sVBRight(prm.sEntree, iLenEntree - 2);
            if (sNumAssert.Length > 0)
            {
                int iPosCom = sNumAssert.IndexOf("'"); // Ignorer le commentaire éventuel
                if (iPosCom > -1) sNumAssert = sNumAssert.Substring(0, iPosCom);
                if (sNumAssert.Length > 0) 
                { 
                    i = clsVBUtil.iVBVal(sNumAssert);
                    i -= 1; // Indice 0 maintenant
                }
            }
        }
        
        if (i < 0 || i >= m_lstBC.Count) 
        {
            string sReponse = sCmdSupprAssertionPb + " " + (i + 1) +
                " <> (" + 1 + ", " + m_lstBC.Count + ")";
            ReponseIAVB(sReponse, bAffiche2Etoiles: true);
            prm.bFinTraitement = true;
            return;
        }

        m_lstBC.RemoveAt(i);

        String s1 = sCmdSupprAssertionOk1 + (i + 1).ToString() + sCmdSupprAssertionOk2;
        ReponseIAVB(s1, bAffiche2Etoiles: true);
        prm.bFinTraitement = true;
        
    }

    private void EffacerBase(TParamEntree prm)
    {
        int iNbAssertions = m_lstBC.Count;
        if (iNbAssertions == 0)
        {
            ReponseIAVB(sErrBaseVide, bAffiche2Etoiles: true);
            prm.bFinTraitement = true;
            return;
        }
        
        ReponseIAVB(sCmdEffacerConfirm, bAffiche2Etoiles: true);
        if (!bQuestionResponseOui(sCmdEffacerConfirm))
        {
            prm.bFinTraitement = true;
            return;
        }
        m_lstBC = new List <TAssertion>();        
        ReponseIAVB(sCmdEffacerOk, bAffiche2Etoiles: true);
        prm.bFinTraitement = true;
    }

    private bool bAssertCMot(TParamEntree prm, int K)
    {
        // Module de recherche d'un prochain aiNumAssertCMot(K)

        int iDeb = prm.aiNumAssertCMot[K];
        int m_iNbAssertions = m_lstBC.Count;
        for (int i = iDeb + 1; i < m_iNbAssertions; i++)
        {
            for (int j = 1; j <= m_lstBC[iDeb].iNbMots; j++)
            {
                if (m_lstBC[i].asMot[j] == prm.asMotsEntree[K])
                {
                    prm.aiNumAssertCMot[K] = i;
                    // Ex.:ANIMAL MANGEUR ?
                    // Réponse : CHAT, TIGRE
                    // asMotsEntree(K) = MANGEUR dans les 2 cas
                    return true;
                }
            }
        }
        return false;
    }

    private bool bComposFonction(TParamEntree prm)
    {
        // Composition de fonction (relation verticale)

        bool bComposFonction = false;
        string[] asMotReduit = new string[iNbMotsEntreeMax+1];

        for (int k = 1; k <= prm.iNbMotsEntree; k++) asMotReduit[k] = prm.asMotsEntree[k];

        int m_iNbAssertions = m_lstBC.Count;
        for (int k = prm.iNbMotsEntree; k >= 2; k--)
        {
            int i;
            string sGdTerme = asMotReduit[k - 1] + asMotReduit[k];

            // Recherche des gd termes sur les mots n° 2 et 3 seulement de la BC :
            //  du type Y = F(X) : Y:1, F:2 et X:3, par ex.: BLANC = COULEUR(ARTABAN)
            bool bCompFct = false;
            int iDeb = m_iNumAssertionEnCours;
            if (m_bVersionModifiee) iDeb = prm.iNumAssertMinRech;
            for (i = iDeb; i < m_iNbAssertions; i++)
            {
                if ((m_lstBC[i].asMot[2] + m_lstBC[i].asMot[3]) == sGdTerme)
                {
                    bCompFct = true;
                    break;
                }
            }
            if (!bCompFct) return bComposFonction;
            asMotReduit[k - 1] = m_lstBC[i].asMot[1];
        }

        // Fin de composition de fct : réponse trouvée
        ReponseIAVB(asMotReduit[1]);
        return true;
        // Ex.:
        // ARTABAN = CHEVAL(HENRI_IV)
        // BLANC = COULEUR(ARTABAN)
        // HENRI_IV = ROI(NAVARRE)
        // COULEUR(CHEVAL(ROI(NAVARRE))) ?
        // Réponse : BLANC
        // sGdTerme = COULEURARTABAN
    }

    private bool bIndirection(TParamEntree prm)
    {
        int iMemNumAssertCMot2 = prm.aiNumAssertCMot[2];
        bool bIndirection = false;
        do
        {
            prm.aiNumAssertCMot[2] = iMemNumAssertCMot2;
            do
            {
                int i1 = prm.aiNumAssertCMot[1];
                int i2 = prm.aiNumAssertCMot[2];
                for (int j1 = 1; j1 <= m_lstBC[i1].iNbMots; j1++)
                {
                    string sGdTerme = m_lstBC[i1].asMot[j1];
                    for (int j2 = 1; j2 <= m_lstBC[i2].iNbMots; j2++)
                    {
                        if (sGdTerme == m_lstBC[i2].asMot[j2] &&
                            i1 != i2 &&
                            sGdTerme != prm.asMotsEntree[1] &&
                            sGdTerme != prm.asMotsEntree[2] &&
                            (prm.iNbMotsEntree <= 3 || !m_bVersionModifiee))
                        {
                            // Modification : il ne peut y avoir indirection que sur une
                            //  question simple (au plus 3 termes signifiants)
                            bIndirection = true;
                            ReponseIAVB(m_lstBC[i1].asMot[j1], bListe:true);
                            j1 = iNbMotsBCMax;
                            j2 = iNbMotsBCMax;

                            // Ex.: QUELLE EST LA COULEUR DU STYLO DE FRANÇOIS ?
                            // Réponse : BLEU
                            // sGdTerme = BLEU, asMotsEntree(1) = COULEUR, asMotsEntree(2) = STYLO

                            // Ex.: QUELLE FILLE EST SAGE ?
                            // Réponse : ANNIE
                            // sGdTerme = ANNIE, asMotsEntree(1) = FILLE, asMotsEntree(2) = SAGE

                            // ARTABAN = CHEVAL(HENRI_IV)
                            // BLANC = COULEUR(ARTABAN)
                            // HENRI_IV = ROI(NAVARRE)
                            // Bonne indirection:
                            // QUEL ROI A UN CHEVAL ? : HENRI_IV
                            // Faille :  mauvaise indirection
                            // CHEVAL DU ROI ? : HENRI_IV
                            // QUELLE EST LA COULEUR DU CHEVAL DU ROI ? : ARTABAN
                            // sGdTerme = ARTABAN
                            // prm.iNbMotsEntree = 3 : COULEUR, CHEVAL, ROI
                            // m_aBC(I1).asMot(J1) = ARTABAN
                            // prm.asMotsEntree(1) = COULEUR
                            // prm.asMotsEntree(2) = CHEVAL
                        }
                    }
                }
            } while (bAssertCMot(prm, 2));
        } while (bAssertCMot(prm, 1));
        ReponseIAVB("", bFinListe:true);
        return bIndirection;
    }

    private static bool bQuestionResponseOui(string sQuestion)
    {
        // ToDo : Confirmation de l'effacement de la base : générer un événement
        //if (MessageBox.Show(sQuestion, sTitreMsg, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
        //{
        //    return false;
        //}
        return true;
    }

    private bool bRelationHorizontale(TParamEntree prm)
    {
        int[] aiMotsTrouves = new int[3];
        bool bRelationHorizontale = false;
        int iDeb = m_iNumAssertionEnCours;
        if (m_bVersionModifiee) iDeb = prm.iNumAssertMinRech;
        int iNbAssertions = m_lstBC.Count;
        for (int i = iDeb; i < iNbAssertions; i++)
        {
            int iNbMotsTrouves = 0;
            int iMemJ = 1;
            for (int k = 1; k <= prm.iNbMotsEntree; k++)
            {
                bool bMotTrouve = false;
                int j = iMemJ;
                while (j <= m_lstBC[i].iNbMots)
                {
                    if (m_lstBC[i].asMot[j] == prm.asMotsEntree[k])
                    {
                        bMotTrouve = true;
                        break;
                    }
                    j++;
                }
                if (bMotTrouve)
                {
                    iNbMotsTrouves++;
                    if (iNbMotsTrouves <= 2) aiMotsTrouves[iNbMotsTrouves] = j;
                }
                iMemJ = j; // Force l'ordre de recherche à Sujet/Verbe/Complement seulement
                // Cela permet de répondre Oui à la question :
                // SABINE AIME-T-ELLE JACQUES ?
                // et cela empêche de trouver une relation H à la question :
                // JACQUES AIME-T-IL SABINE ?
            }

            if (iNbMotsTrouves != 0)
            {
                if (iNbMotsTrouves == prm.iNbMotsEntree && 
                    iNbMotsTrouves == m_lstBC[i].iNbMots)
                {
                    ReponseIAVB(sReponseOui, bAffiche2Etoiles: true);
                    return true;
                }
                if (iNbMotsTrouves == prm.iNbMotsEntree)
                {
                    bRelationHorizontale = true;
                    if (aiMotsTrouves[1] == 1 && aiMotsTrouves[2] == 2)
                    {
                        int iLenE = clsVBUtil.iVBLen(m_lstBC[i].sEntree);
                        String s1 = clsVBUtil.sVBRight(m_lstBC[i].sEntree,
                            iLenE - m_lstBC[i].iPosFinMot2);
                        ReponseIAVB(s1, bListe:true);
                        // Ex.:
                        // QUI MARIE REGARDE-T-ELLE ?
                        // Réponse : HUGUES
                        // aBC(I).sEntree = MARIE REGARDE HUGUES
                    }
                    else if (aiMotsTrouves[1] == 2)
                    {
                        String s1 = clsVBUtil.sVBLeft(m_lstBC[i].sEntree, m_lstBC[i].iPosFinMot1);
                        ReponseIAVB(s1, bListe:true);
                        // Ex.:
                        // CHEF DE SERVICE ?
                        // M.RENE
                        // M.DUBOIS
                        // aBC(I).sEntree = M.RENE EST LE CHEF DU SERVICE COMPTABILITE
                        // aBC(I).sEntree = M.DUBOIS EST CHEF DU SERVICE PHOTO
                    }
                    else if (aiMotsTrouves[1] == 3)
                    {
                        String s1 = clsVBUtil.sVBLeft(m_lstBC[i].sEntree, m_lstBC[i].iPosFinMot2);
                        ReponseIAVB(s1, bListe:true);
                        // Ex.:
                        // SERVICE PHOTO ?
                        // M.DUBOIS EST CHEF
                    }
                    else ReponseIAVB(m_lstBC[i].sEntree, bListe: true); // Liste
                        // Pas d'exemple !
                }
            }
        }
        ReponseIAVB("", bFinListe:true);
        return bRelationHorizontale;
    }

    private void ControleExistenceMots(TParamEntree prm)
    {
        // Interrogation : contrôle de l'existence des mots

        int iMemNumAssertionEnCours = 0;
        int iNbAssertions = m_lstBC.Count;
        m_iNumAssertionEnCours = iNbAssertions-1;
        int iLenEntree = clsVBUtil.iVBLen(prm.sEntree);
        prm.iNumAssertMinRech = iNbAssertions;
        for (int k = 1; k <= prm.iNbMotsEntree; k++)
        {
            bool bMotPresentAssertion = false;
            for (int i = 0; i < iNbAssertions; i++)
            {
                bool bMotTrouve = false;
                int j = 1;
                int iNbMots = m_lstBC[i].iNbMots;
                j = 1;
                while (j <= iNbMots)
                {
                    if (m_lstBC[i].asMot[j] == prm.asMotsEntree[k])
                    { bMotTrouve = true; break; }
                    j++;
                }
                if (bMotTrouve)
                {
                    bMotPresentAssertion = true;
                    if (prm.aiNumAssertCMot[k] == 0) prm.aiNumAssertCMot[k] = i;
                    if (prm.iNbMotsEntree > 1)
                    {
                        iMemNumAssertionEnCours = i;
                        break;
                    }
                    if (j == 2 && iLenEntree > clsVBUtil.iVBLen(prm.asMotsEntree[k]) + 1)
                    {
                        String s1 = clsVBUtil.sVBLeft(m_lstBC[i].sEntree, m_lstBC[i].iPosFinMot1);
                        ReponseIAVB(s1, bListe:true);
                        // Question spécifique, ex.:
                        // ET LAQUELLE EST FOFOLLE ?
                        // MINNA = Left(aBC(I).sEntree, aBC(I).iPosFinMot1)
                        // aBC(I).sEntree = MINNA EST FOFOLLE
                    }
                    else
                    {
                        String s1 = m_lstBC[i].sEntree;
                        ReponseIAVB(s1, bListe:true);
                        // Entrée contenant le mot K : Question générique, ex.:
                        // RESPONSABLE ?
                        // M.BERTRAND EST RESPONSABLE DE L'ANNEXE
                        // M.JACQUES EST RESPONSABLE DE LA SAISIE
                    }
                }
            }
            if (!bMotPresentAssertion)
            {
                string sReponse = sReponseConnaisPas + prm.asMotsEntree[k];
                ReponseIAVB(sReponse, bAffiche2Etoiles: true);
                prm.bFinTraitement = true;
                return; // Pas de liste ici, quitter directement
            }
            if (prm.iNbMotsEntree == 1)
            {
                prm.bFinTraitement = true;
                break; // Liste possible ici
            }

            // Conserver l'indice minimum pour l'espace de recherche max. de l'assertion
            if (m_bVersionModifiee)
            {
                if (iMemNumAssertionEnCours < prm.iNumAssertMinRech)
                    prm.iNumAssertMinRech = iMemNumAssertionEnCours;
            }
            else if (iMemNumAssertionEnCours < m_iNumAssertionEnCours)
                m_iNumAssertionEnCours = iMemNumAssertionEnCours;
        }
        ReponseIAVB("", bFinListe:true);
    }

    private void Syllogisme(bool bRechecheApprofondie)
    {
        // Recherche du moyen terme
        // sPetitTerme : Petit terme extrait de l'assertion pour syllogisme
        // sGdTerme : Concaténation de 2 ou 3 mots signifiants et extrait de l'assertion
        
        bool bPasse2 = false;
        
        int iNbAssertions = m_lstBC.Count;
        int iNumAss2 = iNbAssertions-1;
        if (m_bVersionModifiee) iNumAss2 = m_iNumAssertionEnCours; // Màj si assertion déjà connue
        bool bAucunMotTrouve = true;
        int iAssertionPreced = iNumAss2 - 1;
        if (bRechecheApprofondie) iAssertionPreced = 0; // Si bVersionModifiee seulement
        int iPas = -1;
        int iDebut = iNumAss2 - 1;
        int iFin = iAssertionPreced;

    Passe2:
        if (bPasse2)
        {
            iPas = 1;
            iDebut = iNumAss2 + 1;
            iFin = iNbAssertions-1;
        }
        //for (int iNumAss1 = iDebut; iNumAss1<=iFin; iNumAss1+=iPas)
        //for (int iNumAss1 = iDebut; iNumAss1>=iFin; iNumAss1-=iPas)
        for (int iNumAss1 = iDebut; 
            ((iPas >> 0x1f) ^ iNumAss1) <= ((iPas >> 0x1f) ^ iFin); iNumAss1 += iPas)
        {
            int iPosMT_PMin = 0; // Place du moyen terme dans la prémisse mineure
            int iPosMT_PMaj = 0; // Place du moyen terme dans la prémisse majeure
            
            bool bMotTrouve = false;
            int iLenA1 = clsVBUtil.iVBLen(m_lstBC[iNumAss1].sEntree);
            int iLenA2 = clsVBUtil.iVBLen(m_lstBC[iNumAss2].sEntree);
            string sMotPivot = "";
            int k = 1;
            do
            {
                int j = 1;
                do
                {
                    if (!string.IsNullOrEmpty(m_lstBC[iNumAss1].asMot[j]) &&
                        m_lstBC[iNumAss2].asMot[k] == m_lstBC[iNumAss1].asMot[j])
                    {
                        sMotPivot = m_lstBC[iNumAss2].asMot[k];
                        bMotTrouve = true;
                        bAucunMotTrouve = false;
                        iPosMT_PMaj = j;
                        iPosMT_PMin = k;
                        j = iNbMotsBCMax;
                        k = iNbMotsBCMax;
                    }
                    j++;
                }
                while (j <= 3); 
                k++;
            }
            while (k <= 3);

            if (bRechecheApprofondie && !bMotTrouve) continue;
            if (!bRechecheApprofondie && !bMotTrouve)
            {
                ReponseIAVB(sReponseRienConclure, bAffiche2Etoiles: true);
                return;
            }

            // Résolution
            string sPetitTerme = "";
            bool bInversion = false;
            bool bFctPetitTerme = false;
            bool bFctGdTerme = false;
            switch (iPosMT_PMin)
            {
                case 1:
                    sPetitTerme = clsVBUtil.sVBRight(m_lstBC[iNumAss2].sEntree, iLenA2 - 
                        m_lstBC[iNumAss2].iPosFinMot1);
                    // Ex.: EST PHILOSOPHE
                    if (m_bVersionModifiee && clsVBUtil.sVBLeft(sPetitTerme, 4) == sMotEst + sEspace)
                        bInversion = true;
                    break;

                case 2:
                    sPetitTerme = clsVBUtil.sVBLeft(m_lstBC[iNumAss2].sEntree, 
                        m_lstBC[iNumAss2].iPosFinMot1);
                    // Ex.: "OR SOCRATE "
                    break;

                case 3:
                    sPetitTerme = clsVBUtil.sVBLeft(m_lstBC[iNumAss2].sEntree, 
                        m_lstBC[iNumAss2].iPosFinMot2);
                    if (m_bVersionModifiee && clsVBUtil.sVBRight(sPetitTerme, 1) == sParenthOuv)
                        bFctPetitTerme = true;
                    break;
            }

            // Ex.: OR TOUT HOMME SENSE
            if (clsVBUtil.sVBLeft(sPetitTerme, 3) == sMotOr + sEspace)
                sPetitTerme = clsVBUtil.sVBRight(sPetitTerme, clsVBUtil.iVBLen(sPetitTerme) - 3);

            string sGdTerme="";
            switch (iPosMT_PMaj)
            {
                case 1:
                    sGdTerme = clsVBUtil.sVBRight(m_lstBC[iNumAss1].sEntree, iLenA1 -
                        m_lstBC[iNumAss1].iPosFinMot1);
                    // Ex.: "EST MORTEL"
                    // Ex.: "EST GREC"
                    // Ex.: "EST INCOMPRIS"
                    if (bFctPetitTerme && clsVBUtil.sVBLeft(sGdTerme, 4) == sMotEst + sEspace)
                        sGdTerme = clsVBUtil.sVBTrim(clsVBUtil.sVBRight(sGdTerme, 
                            clsVBUtil.iVBLen(sGdTerme) - 4)) + sEspace + sParenthFerm;
                    break;

                case 2:
                    Syllogisme2(iNumAss1, bFctPetitTerme, bInversion, ref sGdTerme);
                    break;

                case 3:
                    sGdTerme = clsVBUtil.sVBLeft(m_lstBC[iNumAss1].sEntree, 
                        m_lstBC[iNumAss1].iPosFinMot2);
                    if (m_bVersionModifiee && clsVBUtil.sVBRight(sGdTerme, 1) == sParenthOuv)
                        bFctGdTerme = true;
                    break;
            }
            SyllogismeConclusion(iPosMT_PMin, iPosMT_PMaj, sPetitTerme, sGdTerme,
                ref bInversion, bFctGdTerme, sMotPivot);
        }

        if (bRechecheApprofondie && !bPasse2)
        {
            bPasse2 = true;
            goto Passe2;
        }
        ReponseIAVB("", bFinListe:true);
        if (bAucunMotTrouve) ReponseIAVB(sReponseRienConclure, bAffiche2Etoiles: true);
    }

    private void Syllogisme2(int iNumAss1, bool bFctPetitTerme, bool bInversion,
        ref string sGdTerme)
    {
        sGdTerme = clsVBUtil.sVBLeft(m_lstBC[iNumAss1].sEntree, m_lstBC[iNumAss1].iPosFinMot1);

        if (!m_bVersionModifiee) return;

        if (clsVBUtil.sVBLeft(sGdTerme, 3) == sMotOr + sEspace)
        {
            if (bFctPetitTerme)
            {
                sGdTerme = clsVBUtil.sVBTrim(clsVBUtil.sVBRight(sGdTerme,
                    clsVBUtil.iVBLen(sGdTerme) - 3)) + sEspace + sParenthFerm;
                return;
            }

            if (bInversion)
                sGdTerme = clsVBUtil.sVBTrim(clsVBUtil.sVBRight(sGdTerme,
                    clsVBUtil.iVBLen(sGdTerme) - 3));
            else
                sGdTerme = sMotEstComme + sEspace +
                    clsVBUtil.sVBTrim(clsVBUtil.sVBRight(sGdTerme, 
                        clsVBUtil.iVBLen(sGdTerme) - 3));
            return;
        }

        if (bFctPetitTerme)
        { sGdTerme = clsVBUtil.sVBTrim(sGdTerme) + sEspace + sParenthFerm; return; }

        if (clsVBUtil.sVBLeft(sGdTerme, 3) != sMotEst &&
            clsVBUtil.sVBLeft(sGdTerme, 4) != sMotSont &&
            (clsVBUtil.sVBLeft(sGdTerme, 3) != sMotTou || !bInversion))
            sGdTerme = sMotEstComme + sEspace + clsVBUtil.sVBTrim(sGdTerme);
    }

    private void SyllogismeConclusion(int iPosMT_PMin, int iPosMT_PMaj, string sPetitTerme,
        string sGdTerme, ref bool bInversion, bool bFctGdTerme, string sMotPivot)
    {
        if ((iPosMT_PMin + iPosMT_PMaj) == 2)
        {
            if (clsVBUtil.sVBLeft(sPetitTerme, 4) == (sMotEst + sEspace))
            {
                sPetitTerme = clsVBUtil.sVBRight(sPetitTerme, clsVBUtil.iVBLen(sPetitTerme) - 4);
                if (clsVBUtil.sVBLeft(sPetitTerme, 3) == sMotUn + sEspace)
                    sPetitTerme = clsVBUtil.sVBRight(sPetitTerme, clsVBUtil.iVBLen(sPetitTerme) - 3);
                sPetitTerme = sMotQuelque + sEspace + clsVBUtil.sVBTrim(sPetitTerme) + sEspace;
                bInversion = false;
            }
            else if (clsVBUtil.sVBLeft(sPetitTerme, 5) == sMotSont + sEspace)
            {
                sPetitTerme = clsVBUtil.sVBRight(sPetitTerme, clsVBUtil.iVBLen(sPetitTerme) - 5);
                sPetitTerme = sMotQuelques + sEspace + clsVBUtil.sVBTrim(sPetitTerme);
                bInversion = false;
            }
        }
        sPetitTerme = clsVBUtil.sVBTrim(sPetitTerme);
        sGdTerme = clsVBUtil.sVBTrim(sGdTerme);

        if (m_bVersionModifiee)
        {
            if (bInversion)
            {
                if (clsVBUtil.sVBLeft(sGdTerme, 1) == sEgal)
                    sGdTerme = clsVBUtil.sVBMid2(sGdTerme, 2);
            }
            else if (clsVBUtil.sVBLeft(sPetitTerme, 1) == sEgal)
                sPetitTerme = clsVBUtil.sVBMid2(sPetitTerme, 2);
        }

        string sConclusion = sMotDonc + sEspace + sPetitTerme + sEspace + sGdTerme;
        if (bInversion) sConclusion = sMotDonc + sEspace + sGdTerme + sEspace + sPetitTerme;
        if (bFctGdTerme)
        {
            if (bInversion)
                sPetitTerme = clsVBUtil.sVBRight(sPetitTerme, clsVBUtil.iVBLen(sPetitTerme) - 4);
            sConclusion = sMotDonc + sEspace + sGdTerme + sEspace + sPetitTerme +
                sEspace + sParenthFerm;
        }
        string sConclusionParlee = sConclusion;
        if (m_bVersionModifiee) sConclusion = sConclusion + "  <" + sMotPivot + ">";
        ReponseIAVB(sConclusion, sTexteParleSpecifique: sConclusionParlee, bListe: true);

        // Ex. de syllogisme :
        // TOUT HOMME EST MORTEL
        // OR SOCRATE EST UN HOMME
        // DONC SOCRATE EST MORTEL

        // Il y a plusieurs formes de syllogisme, la conclusion peut en varier :
        // PLATON EST GREC
        // OR PLATON EST PHILOSOPHE
        // DONC QUELQUE PHILOSOPHE EST GREC

        // Dans certains cas, un syllogisme peut être trouvé avec 2 mots au lieu d'un seul :
        // TOUT LOGICIEN EST INCOMPRIS
        // OR TOUT HOMME SENSE EST LOGICIEN
        // DONC TOUT HOMME SENSE EST INCOMPRIS
    }

    private void ReponseIAVB(string sTexte, bool bAffiche2Etoiles = false,
        bool bRappelReponse = false, bool bSilence = false, string sTexteParleSpecifique = "",
        bool bListe = false, bool bFinListe = false)
    {
        if (bFinListe)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string sReponse in m_lstReponses)
            {
                if (m_bNormalisationSortieTrimEtVbLf)
                    sb.Append(sReponse.Trim() + vbLf); // 09/04/2017 Normalisation des réponses
                else
                    sb.Append(sReponse + vbLf);
            }
            m_sReponse = sb.ToString();
            sb = new StringBuilder();
            foreach (string sReponse in m_lstReponsesVocales) sb.Append(sReponse + vbLf);
            m_sReponseVocale = sb.ToString();
        }
        else
        {
            string sTexteEcrit = sTexte;
            if (bAffiche2Etoiles) sTexteEcrit = "** " + sTexte;
            if (bRappelReponse) m_sRappelQuestion = sTexteEcrit;
            if (m_bVoixActive && !bSilence)
            {
                string sTexteParle = sTexte;
                if (sTexteParleSpecifique.Length > 0) sTexteParle = sTexteParleSpecifique;
                m_lstReponsesVocales.Add(sTexteParle);
                m_lstReponsesVocales.Add(sTexteEcrit);
                StringBuilder sb = new StringBuilder();
                foreach (string sReponse in m_lstReponsesVocales) sb.Append(sReponse + vbLf);
                m_sReponseVocale = sb.ToString();
            }

            // 09/04/2017 Normalisation des réponses
            if (m_bNormalisationSortieTrimEtVbLf) sTexteEcrit += vbLf;

            if (!bRappelReponse)
            {
                if (bSilence) m_sRappelQuestion = sTexteEcrit;
                else if (bListe) m_lstReponses.Add(sTexteEcrit);
                else m_sReponse = sTexteEcrit;
            }
            m_lstReponsesTot.Add(sTexteEcrit);
        }
    }
    
#endregion
}
}

