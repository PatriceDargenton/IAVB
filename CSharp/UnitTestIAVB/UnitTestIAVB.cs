
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IAVB;

namespace IAVB //UnitTestIAVB
{

[TestClass]
public class UnitTestIAVB
{

IAVB.clsIAVB m_oIAVB;

[TestInitialize]
public void init()
{
    m_oIAVB = new IAVB.clsIAVB();
}

[TestMethod]
public void TestIAVBModifie()
{
    m_oIAVB.Initialiser(bVersionModifiee: true, bNormalisationSortieTrimEtVbLf: false);
    TestNonNormalise();
}
    
[TestMethod]
public void TestIAVBOriginal()
{
    m_oIAVB.Initialiser(bVersionModifiee: false, bNormalisationSortieTrimEtVbLf: false);
    TestNonNormalise();
}

[TestMethod]
public void TestIAVBModifieNorm()
{
    m_oIAVB.Initialiser(bVersionModifiee: true, bNormalisationSortieTrimEtVbLf: true);
    TestNormalise();
}

[TestMethod]
public void TestIAVBOriginalNorm()
{
    m_oIAVB.Initialiser(bVersionModifiee: false, bNormalisationSortieTrimEtVbLf: true);
    TestNormalise();
}

public void TestNormalise()
{
    // Sorties normalisées (toujours Trim et vbLf)

    Assert.AreEqual(clsIAVB.bTraiterEnMinuscules, true);

    const string vbLf = clsIAVB.vbLf; // "\n";
    const string sOK = "** Compris." + vbLf;
    const string sOui = "** Oui." + vbLf;
    const string sNon = "** Non." + vbLf;
    const string sIgnore = "** Je l'ignore." + vbLf;
    const string sDC = "** Assertion déjà connue !" + vbLf;
    const string sRC = "** Je ne peux rien conclure !" + vbLf;

    m_oIAVB.IAVBMain("TOUT STYLO EST BLEU");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);

    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE UN STYLO");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("BLEU EST UNE COULEUR");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ROUGE EST UNE COULEUR");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("TOUT STYLO EST EN PLASTIQUE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LE PLASTIQUE EST UNE MATIERE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE-T-IL UN STYLO ROUGE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU STYLO DE FRANÇOIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    m_oIAVB.IAVBMain("RAOUL A ACHETE UN STYLO EGALEMENT");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DE QUELLE COULEUR EST LE STYLO DE RAOUL ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    m_oIAVB.IAVBMain("TOUT STYLO EST EN QUELLE MATIERE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "plastique" + vbLf);
    m_oIAVB.IAVBMain("EN QUELLE MATIERE EST TOUT STYLO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "plastique" + vbLf);
    m_oIAVB.IAVBMain("' Limitation actuelle du logiciel : réponse fausse :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE-T-IL UN STYLO BLEU ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE-T-IL UN STYLO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    m_oIAVB.IAVBMain("COULEUR STYLO FRANÇOIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    m_oIAVB.IAVBMain("DE QUELLE COULEUR EST LE STYLO DE FRANÇOIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    
    m_oIAVB.IAVBMain("ANNIE EST UNE JOLIE FILLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ANNIE EST SAGE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MINNA EST UNE FILLE ELLE AUSSI");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MINNA EST FOFOLLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUELLE FILLE EST SAGE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "annie" + vbLf);
    m_oIAVB.IAVBMain("ET LAQUELLE EST FOFOLLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "minna" + vbLf);
    m_oIAVB.IAVBMain("JOLIE EST LE CONTRAIRE DE LAIDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("EST-CE QU'ANNIE EST LAIDE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jolie" + vbLf);
    m_oIAVB.IAVBMain("ÇA SIGNIFIE-T-IL QU'ANNIE EST UNE JOLIE FILLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    
    m_oIAVB.IAVBMain("JEAN REGARDE MARIE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MARIE REGARDE HUGUES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI MARIE REGARDE-T-ELLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "hugues" + vbLf);
    m_oIAVB.IAVBMain("MARIE REGARDE-T-ELLE JEAN ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("HUGUES EST LE FRERE D'HENRI");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("HENRI EST LE FILS D'OCTAVE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OCTAVE EST L'ONCLE D'ANATOLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI EST LE FILS DE L'ONCLE D'ANATOLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri" + vbLf);
    m_oIAVB.IAVBMain("ET QUI REGARDE LE FRERE DU FILS DE L'ONCLE D'ANATOLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "marie" + vbLf);
    
    m_oIAVB.IAVBMain("L'ENTREPRISE A UN SIEGE ET UNE ANNEXE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.BERTRAND EST RESPONSABLE DE L'ANNEXE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("L'ANNEXE A 15 SERVICES DIFFERENTS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.JACQUES EST RESPONSABLE DE LA SAISIE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.RENE EST LE CHEF DU SERVICE COMPTABILITE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.MARTIN EST UN AMI DE M.JACQUES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.DUBOIS EST CHEF DU SERVICE PHOTO");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LA SAISIE EST UN SERVICE DECENTRALISE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("RENAUD EST LE FILS DE M.BERTRAND");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DAMIEN EST LE FILS DE M.RENE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("CHEF DE SERVICE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.rene" + vbLf + "m.dubois" + vbLf);
    m_oIAVB.IAVBMain("CHEF DE SERVICE PHOTO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.dubois" + vbLf);
    m_oIAVB.IAVBMain("SERVICE PHOTO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.dubois est chef" + vbLf);
    m_oIAVB.IAVBMain("RESPONSABLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, 
        "m.bertrand est responsable de l'annexe" + vbLf +
        "m.jacques est responsable de la saisie" + vbLf);
    m_oIAVB.IAVBMain("M.MARTIN EST L'AMI DE QUI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "de m.jacques" + vbLf);
    m_oIAVB.IAVBMain("QUI EST LE FILS DU RESPONSABLE DES 15 SERVICES ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "renaud" + vbLf);
    m_oIAVB.IAVBMain("QUI EST LE FILS D'UN CHEF DE SERVICE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "damien" + vbLf);
    m_oIAVB.IAVBMain("QUI EST L'AMI DU RESPONSABLE D'UN SERVICE DECENTRALISE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.martin" + vbLf);
    
    m_oIAVB.IAVBMain("TOUT HOMME EST MORTEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR SOCRATE EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc socrate est mortel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc socrate est mortel" + vbLf);
    
    m_oIAVB.IAVBMain("TOUT HOMME EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR PAUL EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "or paul est un homme :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc paul est bipede  <homme>" + vbLf +
            "donc paul est comme socrate  <homme>" + vbLf +
            "donc paul est mortel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede" + vbLf);
    
    m_oIAVB.IAVBMain("TOUT CE QUI EST RARE EST CHER");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("UN CHEVAL BON.MARCHE EST RARE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc un cheval bon.marche est cher  <rare>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc un cheval bon.marche est cher" + vbLf);
    
    m_oIAVB.IAVBMain("PLATON EST GREC");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR PLATON EST PHILOSOPHE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc quelque philosophe est grec  <platon>" +vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc quelque philosophe est grec" + vbLf);
    
    m_oIAVB.IAVBMain("QUEL PHILOSOPHE EST GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("Y A-T-IL UN PHILOSOPHE GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("QUI EST GREC ET PHILOSOPHE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("PHILOSOPHE GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("PLATON EST-IL GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    
    m_oIAVB.IAVBMain("LES SAVANTS SONT SOUVENT DISTRAITS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR TOUS LES SAVANTS SONT BAVARDS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, 
            "donc quelques bavards sont souvent distraits  <savants>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, 
            "donc quelques bavards sont souvent distraits" + vbLf);
    
    m_oIAVB.IAVBMain("TOUS LES HOMMES SONT DES MORTELS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR DES HOMMES SONT JUSTES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc quelques justes sont des mortels  <hommes>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc quelques justes sont des mortels" + vbLf);
    
    m_oIAVB.IAVBMain("X1 = X2");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("X2 = X3");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("PAR QUOI X1 = X3 ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "x2" + vbLf);
    
    m_oIAVB.IAVBMain("Y1 = F DE X");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("F DE X ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1" + vbLf);
    m_oIAVB.IAVBMain("X = G DE T1");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("V = K DE W");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("T1 = H DE V");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("F DE G DE H DE K DE W ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1" + vbLf);
    m_oIAVB.IAVBMain("A QUOI EST = F DE G DE H DE K DE W ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1" + vbLf);
    m_oIAVB.IAVBMain("F G H K W ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1" + vbLf);
    
    m_oIAVB.IAVBMain("CHAT = ANIMAL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("CHAT = MANGEUR(SOURIS)");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("TIGRE = ANIMAL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("TIGRE = MANGEUR(HOMME)");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ANIMAL MANGEUR ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "chat" + vbLf + "tigre" + vbLf);
    m_oIAVB.IAVBMain("MANGEUR D'HOMME = ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "tigre" + vbLf);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc  mangeur(homme) = animal  <tigre>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc = mangeur(homme) = animal" + vbLf);
        //Assert.AreEqual(m_oIAVB.m_sReponse, "donc  mangeur(homme) = animal" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "tigre = mangeur(homme) :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc  mangeur(homme) = animal  <tigre>" + vbLf +
            "donc tigre est comme chat  <mangeur>" + vbLf +
            "donc tigre = mangeur( paul )  <homme>" + vbLf +
            "donc tigre = mangeur( bipede )  <homme>" + vbLf +
            "donc tigre = mangeur( socrate )  <homme>" + vbLf +
            "donc tigre = mangeur( mortel )  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc = mangeur(homme) = animal" + vbLf);
        //Assert.AreEqual(m_oIAVB.m_sReponse, "donc  mangeur(homme) = animal" + vbLf);
    
    m_oIAVB.IAVBMain("Y A-T-IL UN MANGEUR DE SOURIS ET D'HOMME ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("QUI EST MANGEUR D'HOMME ET DE SOURIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    
    m_oIAVB.IAVBMain("ARTABAN = CHEVAL DE HENRI.4");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("BLANC = COULEUR D'ARTABAN");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("HENRI.4 = ROI DE NAVARRE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LOUIS.14 = ROI DE FRANCE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("COULEUR DU CHEVAL DU ROI DE NAVARRE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "blanc" + vbLf);
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CHEVAL DU ROI DE NAVARRE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "blanc" + vbLf);
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CHEVAL BLANC DU ROI DE NAVARRE ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "artaban" + vbLf);
    
    m_oIAVB.IAVBMain("' Faille : composition incomplète, mauvaise indirection :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CHEVAL DU ROI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "artaban" + vbLf);
    m_oIAVB.IAVBMain("' Bonne indirection :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("QUEL ROI A UN CHEVAL ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri.4" + vbLf);
    m_oIAVB.IAVBMain("' Faille :  mauvaise indirection :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("CHEVAL DU ROI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri.4" + vbLf);
    m_oIAVB.IAVBMain("CHEVAL DU ROI DE NAVARRE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "artaban" + vbLf);
    m_oIAVB.IAVBMain("CHEVAL DU ROI DE FRANCE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri.4" + vbLf);
    
    m_oIAVB.IAVBMain("LE CANARI EST UN OISEAU JAUNE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("JAUNE EST UNE COULEUR");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUEL OISEAU EST JAUNE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "le canari" + vbLf);
    m_oIAVB.IAVBMain("QUEL EST L'OISEAU JAUNE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "le canari" + vbLf);
    m_oIAVB.IAVBMain("DE QUELLE COULEUR EST LE CANARI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jaune" + vbLf);
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CANARI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jaune" + vbLf);
    m_oIAVB.IAVBMain("COULEUR CANARI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jaune" + vbLf);
    m_oIAVB.IAVBMain("COULEUR ?");
    Assert.AreEqual(m_oIAVB.m_sReponse,
        "bleu est une couleur" + vbLf +
        "rouge est une couleur" + vbLf +
        "blanc = couleur d'artaban" + vbLf +
        "jaune est une couleur" + vbLf);
    
    m_oIAVB.IAVBMain("MARSEILLE EST LA VILLE PHOCEENNE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DEFERRE EST LE MAIRE DE MARSEILLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("PHOCEENNE SIGNIFIE ORIGINAIRE DE PHOCEE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LA PHOCEE EST UNE PROVINCE GRECQUE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("GASTON EST LE PRENOM DE DEFERRE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUEL EST LE PRENOM DU MAIRE DE MARSEILLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "gaston" + vbLf);
    m_oIAVB.IAVBMain("QUEL EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE GRECQUE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "gaston" + vbLf);
    m_oIAVB.IAVBMain("' Test du système : composition de fonctions incomplète :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("QUEL EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "deferre" + vbLf);
    
    m_oIAVB.IAVBMain("' Autres réponses souhaitées : GASTON EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE PHOCEENNE");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("' ou bien : D'UNE PROVINCE PHOCEENNE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("' ou bien : ** heu... GASTON ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    
    m_oIAVB.IAVBMain("PAUL POSSEDE UN PERROQUET BAVARD");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MULTICOLORE SIGNIFIE DE PLUSIEURS COULEURS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("UN PERROQUET EST UN ANIMAL MULTICOLORE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI POSSEDE UN ANIMAL DE PLUSIEURS COULEURS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "paul" + vbLf);
    
    m_oIAVB.IAVBMain("MARIE EST UNE JOLIE FILLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("JOLIE EST LE CONTRAIRE DE LAIDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("MARIE EST-ELLE LAIDE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jolie" + vbLf);
    m_oIAVB.IAVBMain("EST-CE QUE MARIE EST LAIDE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jolie" + vbLf);
    
    m_oIAVB.IAVBMain("SABINE AIME JACQUES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI AIME JACQUES ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "sabine" + vbLf);
    m_oIAVB.IAVBMain("QUI AIME SABINE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    m_oIAVB.IAVBMain("QUI JACQUES AIME-T-IL ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    m_oIAVB.IAVBMain("' Meilleure réponse : Je L'ignore :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("JACQUES AIME-T-IL SABINE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("SABINE AIME-T-ELLE JACQUES ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    
    m_oIAVB.IAVBMain("' Capacités de la version modifiée :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("' Ordre des termes : DONC TOUT CHEVAL EST HERBIVORE :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT CHEVAL EST UN EQUIDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR TOUT EQUIDE EST HERBIVORE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc tout cheval est herbivore  <equide>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc est herbivore tout cheval" + vbLf);
    
    m_oIAVB.IAVBMain("' Sens logique : DONC PAUL EST COMME TOUT HOMME :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT HOMME EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("OR PAUL EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sRC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "or paul est bipede :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc quelque bipede possede un perroquet bavard  <paul>" + vbLf +
            "donc quelque bipede est un homme  <paul>" + vbLf +
            "donc paul est comme tout homme  <bipede>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, sRC);
    
    m_oIAVB.IAVBMain("' Sens logique : ARISTOTE EST COMME TOUT HOMME :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT HOMME EST RATIONNEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR ARISTOTE EST RATIONNEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc aristote est comme tout homme  <rationnel>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc aristote tout homme" + vbLf);
    
    m_oIAVB.IAVBMain("ARISTOTE ÉTAIT GREC");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ARISTOTE ÉTAIT PHILOSOPHE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("PHILOSOPHE GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf + "aristote" + vbLf);
    
    m_oIAVB.IAVBMain("' Syllogismes : ordre des assertions");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("OR SOCRATE EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("TOUT HOMME EST MORTEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, sRC);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "tout homme est mortel :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc socrate est mortel  <homme>" + vbLf +
            "donc quelque mortel est bipede  <homme>" + vbLf +
            "donc paul est mortel  <homme>" + vbLf +
            "donc tigre = mangeur( mortel )  <homme>" + vbLf +
            "donc quelque mortel est rationnel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("' Syllogismes : plusieurs à la fois");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT HOMME EST MORTEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("TOUT HOMME EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("OR SOCRATE EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc socrate est mortel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("OR PAUL EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
        "or paul est un homme :" + vbLf +
        "** Recherche approfondie..." + vbLf +
        "donc paul est bipede  <homme>" + vbLf +
        "donc paul est comme socrate  <homme>" + vbLf +
        "donc paul est mortel  <homme>" + vbLf +
        "donc tigre = mangeur( paul )  <homme>" + vbLf +
        "donc quelque homme possede un perroquet bavard  <paul>" + vbLf +
        "donc quelque homme est bipede  <paul>" + vbLf +
        "donc paul est rationnel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("TOUT LOGICIEN EST INCOMPRIS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR TOUT HOMME SENSE EST LOGICIEN");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc tout homme sense est incompris  <logicien>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc tout homme sense est incompris" + vbLf);
    
    m_oIAVB.IAVBMain("' Syllogismes : syntaxe imparfaite (de la version modifiée)");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("LE LOUVRE EST BEAU");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR J'AIME TOUT CE QUI EST BEAU");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc j'aime est comme le louvre  <beau>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc j'aime le louvre" + vbLf);
}

public void TestNonNormalise()
{
    // Sorties non normalisées (parfois Trim et vbLf)
    
    Assert.AreEqual(clsIAVB.bTraiterEnMinuscules, true);

    const string vbLf = clsIAVB.vbLf; // "\n";
    const string sOK = "** Compris.";
    const string sOui = "** Oui.";
    const string sNon = "** Non.";
    const string sIgnore = "** Je l'ignore.";
    const string sDC = "** Assertion déjà connue !";
    const string sRC = "** Je ne peux rien conclure !";

    m_oIAVB.IAVBMain("TOUT STYLO EST BLEU");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);

    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE UN STYLO");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("BLEU EST UNE COULEUR");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ROUGE EST UNE COULEUR");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("TOUT STYLO EST EN PLASTIQUE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LE PLASTIQUE EST UNE MATIERE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE-T-IL UN STYLO ROUGE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU STYLO DE FRANÇOIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    m_oIAVB.IAVBMain("RAOUL A ACHETE UN STYLO EGALEMENT");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DE QUELLE COULEUR EST LE STYLO DE RAOUL ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    m_oIAVB.IAVBMain("TOUT STYLO EST EN QUELLE MATIERE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "plastique" + vbLf);
    m_oIAVB.IAVBMain("EN QUELLE MATIERE EST TOUT STYLO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "plastique" + vbLf);
    m_oIAVB.IAVBMain("' Limitation actuelle du logiciel : réponse fausse :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE-T-IL UN STYLO BLEU ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("FRANÇOIS POSSEDE-T-IL UN STYLO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    m_oIAVB.IAVBMain("COULEUR STYLO FRANÇOIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    m_oIAVB.IAVBMain("DE QUELLE COULEUR EST LE STYLO DE FRANÇOIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "bleu" + vbLf);
    
    m_oIAVB.IAVBMain("ANNIE EST UNE JOLIE FILLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ANNIE EST SAGE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MINNA EST UNE FILLE ELLE AUSSI");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MINNA EST FOFOLLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUELLE FILLE EST SAGE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "annie" + vbLf);
    m_oIAVB.IAVBMain("ET LAQUELLE EST FOFOLLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "minna " + vbLf);
    m_oIAVB.IAVBMain("JOLIE EST LE CONTRAIRE DE LAIDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("EST-CE QU'ANNIE EST LAIDE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jolie" + vbLf);
    m_oIAVB.IAVBMain("ÇA SIGNIFIE-T-IL QU'ANNIE EST UNE JOLIE FILLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    
    m_oIAVB.IAVBMain("JEAN REGARDE MARIE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MARIE REGARDE HUGUES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI MARIE REGARDE-T-ELLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "hugues" + vbLf);
    m_oIAVB.IAVBMain("MARIE REGARDE-T-ELLE JEAN ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("HUGUES EST LE FRERE D'HENRI");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("HENRI EST LE FILS D'OCTAVE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OCTAVE EST L'ONCLE D'ANATOLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI EST LE FILS DE L'ONCLE D'ANATOLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri");
    m_oIAVB.IAVBMain("ET QUI REGARDE LE FRERE DU FILS DE L'ONCLE D'ANATOLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "marie");
    
    m_oIAVB.IAVBMain("L'ENTREPRISE A UN SIEGE ET UNE ANNEXE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.BERTRAND EST RESPONSABLE DE L'ANNEXE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("L'ANNEXE A 15 SERVICES DIFFERENTS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.JACQUES EST RESPONSABLE DE LA SAISIE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.RENE EST LE CHEF DU SERVICE COMPTABILITE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.MARTIN EST UN AMI DE M.JACQUES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("M.DUBOIS EST CHEF DU SERVICE PHOTO");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LA SAISIE EST UN SERVICE DECENTRALISE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("RENAUD EST LE FILS DE M.BERTRAND");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DAMIEN EST LE FILS DE M.RENE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("CHEF DE SERVICE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.rene " + vbLf + "m.dubois " + vbLf);
    m_oIAVB.IAVBMain("CHEF DE SERVICE PHOTO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.dubois " + vbLf);
    m_oIAVB.IAVBMain("SERVICE PHOTO ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.dubois est chef " + vbLf);
    m_oIAVB.IAVBMain("RESPONSABLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, 
        "m.bertrand est responsable de l'annexe" + vbLf +
        "m.jacques est responsable de la saisie" + vbLf);
    m_oIAVB.IAVBMain("M.MARTIN EST L'AMI DE QUI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "de m.jacques" + vbLf);
    m_oIAVB.IAVBMain("QUI EST LE FILS DU RESPONSABLE DES 15 SERVICES ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "renaud");
    m_oIAVB.IAVBMain("QUI EST LE FILS D'UN CHEF DE SERVICE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "damien");
    m_oIAVB.IAVBMain("QUI EST L'AMI DU RESPONSABLE D'UN SERVICE DECENTRALISE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "m.martin");
    
    m_oIAVB.IAVBMain("TOUT HOMME EST MORTEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR SOCRATE EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc socrate est mortel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc socrate est mortel" + vbLf);
    
    m_oIAVB.IAVBMain("TOUT HOMME EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR PAUL EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "or paul est un homme :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc paul est bipede  <homme>" + vbLf +
            "donc paul est comme socrate  <homme>" + vbLf +
            "donc paul est mortel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede" + vbLf);
    
    m_oIAVB.IAVBMain("TOUT CE QUI EST RARE EST CHER");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("UN CHEVAL BON.MARCHE EST RARE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc un cheval bon.marche est cher  <rare>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc un cheval bon.marche est cher" + vbLf);
    
    m_oIAVB.IAVBMain("PLATON EST GREC");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR PLATON EST PHILOSOPHE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc quelque philosophe est grec  <platon>" +vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc quelque philosophe est grec" + vbLf);
    
    m_oIAVB.IAVBMain("QUEL PHILOSOPHE EST GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("Y A-T-IL UN PHILOSOPHE GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("QUI EST GREC ET PHILOSOPHE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("PHILOSOPHE GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf);
    m_oIAVB.IAVBMain("PLATON EST-IL GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    
    m_oIAVB.IAVBMain("LES SAVANTS SONT SOUVENT DISTRAITS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR TOUS LES SAVANTS SONT BAVARDS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, 
            "donc quelques bavards sont souvent distraits  <savants>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, 
            "donc quelques bavards sont souvent distraits" + vbLf);
    
    m_oIAVB.IAVBMain("TOUS LES HOMMES SONT DES MORTELS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR DES HOMMES SONT JUSTES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc quelques justes sont des mortels  <hommes>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc quelques justes sont des mortels" + vbLf);
    
    m_oIAVB.IAVBMain("X1 = X2");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("X2 = X3");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("PAR QUOI X1 = X3 ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "x2" + vbLf);
    
    m_oIAVB.IAVBMain("Y1 = F DE X");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("F DE X ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1 " + vbLf);
    m_oIAVB.IAVBMain("X = G DE T1");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("V = K DE W");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("T1 = H DE V");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("F DE G DE H DE K DE W ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1");
    m_oIAVB.IAVBMain("A QUOI EST = F DE G DE H DE K DE W ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1");
    m_oIAVB.IAVBMain("F G H K W ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "y1");
    
    m_oIAVB.IAVBMain("CHAT = ANIMAL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("CHAT = MANGEUR(SOURIS)");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("TIGRE = ANIMAL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("TIGRE = MANGEUR(HOMME)");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ANIMAL MANGEUR ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "chat" + vbLf + "tigre" + vbLf);
    m_oIAVB.IAVBMain("MANGEUR D'HOMME = ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "tigre " + vbLf);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc  mangeur(homme) = animal  <tigre>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc = mangeur(homme) = animal" + vbLf);
        //Assert.AreEqual(m_oIAVB.m_sReponse, "donc  mangeur(homme) = animal" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "tigre = mangeur(homme) :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc  mangeur(homme) = animal  <tigre>" + vbLf +
            "donc tigre est comme chat  <mangeur>" + vbLf +
            "donc tigre = mangeur( paul )  <homme>" + vbLf +
            "donc tigre = mangeur( bipede )  <homme>" + vbLf +
            "donc tigre = mangeur( socrate )  <homme>" + vbLf +
            "donc tigre = mangeur( mortel )  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc = mangeur(homme) = animal" + vbLf);
        //Assert.AreEqual(m_oIAVB.m_sReponse, "donc  mangeur(homme) = animal" + vbLf);
    
    m_oIAVB.IAVBMain("Y A-T-IL UN MANGEUR DE SOURIS ET D'HOMME ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("QUI EST MANGEUR D'HOMME ET DE SOURIS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    
    m_oIAVB.IAVBMain("ARTABAN = CHEVAL DE HENRI.4");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("BLANC = COULEUR D'ARTABAN");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("HENRI.4 = ROI DE NAVARRE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LOUIS.14 = ROI DE FRANCE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("COULEUR DU CHEVAL DU ROI DE NAVARRE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "blanc");
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CHEVAL DU ROI DE NAVARRE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "blanc");
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CHEVAL BLANC DU ROI DE NAVARRE ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "artaban" + vbLf);
    
    m_oIAVB.IAVBMain("' Faille : composition incomplète, mauvaise indirection :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CHEVAL DU ROI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "artaban" + vbLf);
    m_oIAVB.IAVBMain("' Bonne indirection :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("QUEL ROI A UN CHEVAL ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri.4" + vbLf);
    m_oIAVB.IAVBMain("' Faille :  mauvaise indirection :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("CHEVAL DU ROI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri.4" + vbLf);
    m_oIAVB.IAVBMain("CHEVAL DU ROI DE NAVARRE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "artaban");
    m_oIAVB.IAVBMain("CHEVAL DU ROI DE FRANCE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "henri.4" + vbLf);
    
    m_oIAVB.IAVBMain("LE CANARI EST UN OISEAU JAUNE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("JAUNE EST UNE COULEUR");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUEL OISEAU EST JAUNE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "le canari " + vbLf);
    m_oIAVB.IAVBMain("QUEL EST L'OISEAU JAUNE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "le canari " + vbLf);
    m_oIAVB.IAVBMain("DE QUELLE COULEUR EST LE CANARI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jaune" + vbLf);
    m_oIAVB.IAVBMain("QUELLE EST LA COULEUR DU CANARI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jaune" + vbLf);
    m_oIAVB.IAVBMain("COULEUR CANARI ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jaune" + vbLf);
    m_oIAVB.IAVBMain("COULEUR ?");
    Assert.AreEqual(m_oIAVB.m_sReponse,
        "bleu est une couleur" + vbLf +
        "rouge est une couleur" + vbLf +
        "blanc = couleur d'artaban" + vbLf +
        "jaune est une couleur" + vbLf);
    
    m_oIAVB.IAVBMain("MARSEILLE EST LA VILLE PHOCEENNE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DEFERRE EST LE MAIRE DE MARSEILLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("PHOCEENNE SIGNIFIE ORIGINAIRE DE PHOCEE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("LA PHOCEE EST UNE PROVINCE GRECQUE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("GASTON EST LE PRENOM DE DEFERRE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUEL EST LE PRENOM DU MAIRE DE MARSEILLE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "gaston");
    m_oIAVB.IAVBMain("QUEL EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE GRECQUE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "gaston");
    m_oIAVB.IAVBMain("' Test du système : composition de fonctions incomplète :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("QUEL EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "deferre" + vbLf);
    
    m_oIAVB.IAVBMain("' Autres réponses souhaitées : GASTON EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE PHOCEENNE");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("' ou bien : D'UNE PROVINCE PHOCEENNE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("' ou bien : ** heu... GASTON ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    
    m_oIAVB.IAVBMain("PAUL POSSEDE UN PERROQUET BAVARD");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("MULTICOLORE SIGNIFIE DE PLUSIEURS COULEURS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("UN PERROQUET EST UN ANIMAL MULTICOLORE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI POSSEDE UN ANIMAL DE PLUSIEURS COULEURS ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "paul");
    
    m_oIAVB.IAVBMain("MARIE EST UNE JOLIE FILLE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("JOLIE EST LE CONTRAIRE DE LAIDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("MARIE EST-ELLE LAIDE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jolie" + vbLf);
    m_oIAVB.IAVBMain("EST-CE QUE MARIE EST LAIDE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "jolie" + vbLf);
    
    m_oIAVB.IAVBMain("SABINE AIME JACQUES");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("QUI AIME JACQUES ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "sabine " + vbLf);
    m_oIAVB.IAVBMain("QUI AIME SABINE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    m_oIAVB.IAVBMain("QUI JACQUES AIME-T-IL ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sIgnore);
    m_oIAVB.IAVBMain("' Meilleure réponse : Je L'ignore :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("JACQUES AIME-T-IL SABINE ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sNon);
    m_oIAVB.IAVBMain("SABINE AIME-T-ELLE JACQUES ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOui);
    
    m_oIAVB.IAVBMain("' Capacités de la version modifiée :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("' Ordre des termes : DONC TOUT CHEVAL EST HERBIVORE :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT CHEVAL EST UN EQUIDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR TOUT EQUIDE EST HERBIVORE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc tout cheval est herbivore  <equide>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc est herbivore tout cheval" + vbLf);
    
    m_oIAVB.IAVBMain("' Sens logique : DONC PAUL EST COMME TOUT HOMME :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT HOMME EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("OR PAUL EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, sRC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "or paul est bipede :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc quelque bipede possede un perroquet bavard  <paul>" + vbLf +
            "donc quelque bipede est un homme  <paul>" + vbLf +
            "donc paul est comme tout homme  <bipede>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, sRC);
    
    m_oIAVB.IAVBMain("' Sens logique : ARISTOTE EST COMME TOUT HOMME :");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT HOMME EST RATIONNEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR ARISTOTE EST RATIONNEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc aristote est comme tout homme  <rationnel>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc aristote tout homme" + vbLf);
    
    m_oIAVB.IAVBMain("ARISTOTE ÉTAIT GREC");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("ARISTOTE ÉTAIT PHILOSOPHE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("PHILOSOPHE GREC ?");
    Assert.AreEqual(m_oIAVB.m_sReponse, "platon" + vbLf + "aristote" + vbLf);
    
    m_oIAVB.IAVBMain("' Syllogismes : ordre des assertions");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("OR SOCRATE EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("TOUT HOMME EST MORTEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, sRC);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "tout homme est mortel :" + vbLf +
            "** Recherche approfondie..." + vbLf +
            "donc socrate est mortel  <homme>" + vbLf +
            "donc quelque mortel est bipede  <homme>" + vbLf +
            "donc paul est mortel  <homme>" + vbLf +
            "donc tigre = mangeur( mortel )  <homme>" + vbLf +
            "donc quelque mortel est rationnel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("' Syllogismes : plusieurs à la fois");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("TOUT HOMME EST MORTEL");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("TOUT HOMME EST BIPEDE");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("OR SOCRATE EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc socrate est mortel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("OR PAUL EST UN HOMME");
    Assert.AreEqual(m_oIAVB.m_sReponse, sDC);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc paul est bipede  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
        "or paul est un homme :" + vbLf +
        "** Recherche approfondie..." + vbLf +
        "donc paul est bipede  <homme>" + vbLf +
        "donc paul est comme socrate  <homme>" + vbLf +
        "donc paul est mortel  <homme>" + vbLf +
        "donc tigre = mangeur( paul )  <homme>" + vbLf +
        "donc quelque homme possede un perroquet bavard  <paul>" + vbLf +
        "donc quelque homme est bipede  <paul>" + vbLf +
        "donc paul est rationnel  <homme>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc était philosophe était grec" + vbLf);
    
    m_oIAVB.IAVBMain("TOUT LOGICIEN EST INCOMPRIS");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR TOUT HOMME SENSE EST LOGICIEN");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc tout homme sense est incompris  <logicien>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse,
            "donc tout homme sense est incompris" + vbLf);
    
    m_oIAVB.IAVBMain("' Syllogismes : syntaxe imparfaite (de la version modifiée)");
    Assert.AreEqual(m_oIAVB.m_sReponse, "");
    m_oIAVB.IAVBMain("LE LOUVRE EST BEAU");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("OR J'AIME TOUT CE QUI EST BEAU");
    Assert.AreEqual(m_oIAVB.m_sReponse, sOK);
    m_oIAVB.IAVBMain("DONC ?");
    if (m_oIAVB.m_bVersionModifiee)
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc j'aime est comme le louvre  <beau>" + vbLf);
    else
        Assert.AreEqual(m_oIAVB.m_sReponse, "donc j'aime le louvre" + vbLf);
}

}
}
