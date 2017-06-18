
using System.Collections.Generic;

namespace IAVB
{
public class clsInitIAVB {
    
    private List<string> m_lst;
    
    public clsInitIAVB(List<string> lst) {
        m_lst = lst;
    }
    
    private void AjouterExemple(string sExemple) {
        m_lst.Add(sExemple);
    }
    
    public void InitAssertionsExemples() {

        AjouterExemple("");
        AjouterExemple("TOUT STYLO EST BLEU");
        AjouterExemple("FRANÇOIS POSSÈDE UN STYLO");
        AjouterExemple("BLEU EST UNE COULEUR");
        AjouterExemple("ROUGE EST UNE COULEUR");
        AjouterExemple("TOUT STYLO EST EN PLASTIQUE");
        AjouterExemple("LE PLASTIQUE EST UNE MATIÈRE");
        AjouterExemple("FRANÇOIS POSSÈDE-T-IL UN STYLO ROUGE ?");
        AjouterExemple("QUELLE EST LA COULEUR DU STYLO DE FRANÇOIS ?");
        AjouterExemple("RAOUL A ACHETÉ UN STYLO ÉGALEMENT");
        AjouterExemple("DE QUELLE COULEUR EST LE STYLO DE RAOUL ?");
        AjouterExemple("TOUT STYLO EST EN QUELLE MATIÈRE ?");
        AjouterExemple("EN QUELLE MATIÈRE EST TOUT STYLO ?");
        AjouterExemple("' Limitation actuelle du logiciel : réponse fausse :");
        AjouterExemple("FRANÇOIS POSSÈDE-T-IL UN STYLO BLEU ?");
        AjouterExemple("FRANÇOIS POSSÈDE-T-IL UN STYLO ?");
        AjouterExemple("COULEUR STYLO FRANÇOIS ?");
        AjouterExemple("DE QUELLE COULEUR EST LE STYLO DE FRANÇOIS ?");

        AjouterExemple("");
        AjouterExemple("ANNIE EST UNE JOLIE FILLE");
        AjouterExemple("ANNIE EST SAGE");
        AjouterExemple("MINNA EST UNE FILLE ELLE AUSSI");
        AjouterExemple("MINNA EST FOFOLLE");
        AjouterExemple("QUELLE FILLE EST SAGE ?");
        AjouterExemple("ET LAQUELLE EST FOFOLLE ?");
        AjouterExemple("JOLIE EST LE CONTRAIRE DE LAIDE");
        AjouterExemple("EST-CE QU'ANNIE EST LAIDE ?");
        AjouterExemple("ÇA SIGNIFIE-T-IL QU'ANNIE EST UNE JOLIE FILLE ?");

        AjouterExemple("");
        AjouterExemple("JEAN REGARDE MARIE");
        AjouterExemple("MARIE REGARDE HUGUES");
        AjouterExemple("QUI MARIE REGARDE-T-ELLE ?");
        AjouterExemple("MARIE REGARDE-T-ELLE JEAN ?");
        AjouterExemple("HUGUES EST LE FRÈRE D'HENRI");
        AjouterExemple("HENRI EST LE FILS D'OCTAVE");
        AjouterExemple("OCTAVE EST L'ONCLE D'ANATOLE");
        AjouterExemple("QUI EST LE FILS DE L'ONCLE D'ANATOLE ?");
        AjouterExemple("ET QUI REGARDE LE FRÈRE DU FILS DE L'ONCLE D'ANATOLE ?");

        AjouterExemple("");
        AjouterExemple("L'ENTREPRISE A UN SIÈGE ET UNE ANNEXE");
        AjouterExemple("M.BERTRAND EST RESPONSABLE DE L'ANNEXE");
        AjouterExemple("L'ANNEXE A 15 SERVICES DIFFÉRENTS");
        AjouterExemple("M.JACQUES EST RESPONSABLE DE LA SAISIE");
        AjouterExemple("M.RENÉ EST LE CHEF DU SERVICE COMPTABILITÉ");
        AjouterExemple("M.MARTIN EST UN AMI DE M.JACQUES");
        AjouterExemple("M.DUBOIS EST CHEF DU SERVICE PHOTO");
        AjouterExemple("LA SAISIE EST UN SERVICE DÉCENTRALISÉ");
        AjouterExemple("RENAUD EST LE FILS DE M.BERTRAND");
        AjouterExemple("DAMIEN EST LE FILS DE M.RENÉ");
        AjouterExemple("CHEF DE SERVICE ?");
        AjouterExemple("CHEF DE SERVICE PHOTO ?");
        AjouterExemple("SERVICE PHOTO ?");
        AjouterExemple("RESPONSABLE ?");
        AjouterExemple("M.MARTIN EST L'AMI DE QUI ?");
        AjouterExemple("QUI EST LE FILS DU RESPONSABLE DES 15 SERVICES ?");
        AjouterExemple("QUI EST LE FILS D'UN CHEF DE SERVICE ?");
        AjouterExemple("QUI EST L'AMI DU RESPONSABLE D'UN SERVICE DÉCENTRALISÉ ?");

        AjouterExemple("");
        AjouterExemple("TOUT HOMME EST MORTEL");
        AjouterExemple("OR SOCRATE EST UN HOMME");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("TOUT HOMME EST BIPÈDE");
        AjouterExemple("OR PAUL EST UN HOMME");
        AjouterExemple("DONC ?");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("TOUT CE QUI EST RARE EST CHER");
        //AjouterExemple("UN CHEVAL BON_MARCHE EST RARE");
        AjouterExemple("UN CHEVAL BON.MARCHÉ EST RARE");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("PLATON EST GREC");
        AjouterExemple("OR PLATON EST PHILOSOPHE");
        AjouterExemple("DONC ?");
        AjouterExemple("QUEL PHILOSOPHE EST GREC ?");
        AjouterExemple("Y A-T-IL UN PHILOSOPHE GREC ?");
        AjouterExemple("QUI EST GREC ET PHILOSOPHE ?");
        AjouterExemple("PHILOSOPHE GREC ?");
        AjouterExemple("PLATON EST-IL GREC ?");

        AjouterExemple("");
        AjouterExemple("LES SAVANTS SONT SOUVENT DISTRAITS");
        AjouterExemple("OR TOUS LES SAVANTS SONT BAVARDS");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("TOUS LES HOMMES SONT DES MORTELS");
        AjouterExemple("OR DES HOMMES SONT JUSTES");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("X1 = X2");
        AjouterExemple("X2 = X3");
        AjouterExemple("PAR QUOI X1 = X3 ?");

        AjouterExemple("");
        //AjouterExemple("Y1 = F(X)");
        //AjouterExemple("F(X) ?");
        //AjouterExemple("X = G(T1)");
        //AjouterExemple("V = K(W)");
        //AjouterExemple("T1 = H(V)");
        //AjouterExemple("F(G(H(K(W)) ?");
        AjouterExemple("Y1 = F DE X");
        AjouterExemple("F DE X ?");
        AjouterExemple("X = G DE T1");
        AjouterExemple("V = K DE W");
        AjouterExemple("T1 = H DE V");
        AjouterExemple("F DE G DE H DE K DE W ?");
        AjouterExemple("A QUOI EST = F DE G DE H DE K DE W ?");
        AjouterExemple("F G H K W ?");

        AjouterExemple("");
        AjouterExemple("CHAT = ANIMAL");
        AjouterExemple("CHAT = MANGEUR(SOURIS)");
        //AjouterExemple("CHAT = MANGEUR DE SOURIS");
        AjouterExemple("TIGRE = ANIMAL");
        AjouterExemple("TIGRE = MANGEUR(HOMME)");
        // Problème de syntaxe des syllogismes basée sur la détection des ();
        //  si on change, certaines déductions ne marche plus :;
        //AjouterExemple("TIGRE = MANGEUR D'HOMME");
        AjouterExemple("ANIMAL MANGEUR ?");
        //AjouterExemple("MANGEUR(HOMME) = ?");
        AjouterExemple("MANGEUR D'HOMME = ?");
        AjouterExemple("DONC ?");
        AjouterExemple("DONC ?");
        //AjouterExemple("Y A-T-IL UN MANGEUR(SOURIS ET HOMME) ?");
        AjouterExemple("Y A-T-IL UN MANGEUR DE SOURIS ET D'HOMME ?");
        //AjouterExemple("QUI EST MANGEUR(HOMME ET SOURIS) ?");
        AjouterExemple("QUI EST MANGEUR D'HOMME ET DE SOURIS ?");

        AjouterExemple("");
        //AjouterExemple("ARTABAN = CHEVAL(HENRI_IV)");
        AjouterExemple("ARTABAN = CHEVAL DE HENRI.4");
        //AjouterExemple("BLANC = COULEUR(ARTABAN)");
        AjouterExemple("BLANC = COULEUR D'ARTABAN");
        //AjouterExemple("HENRI_IV = ROI(NAVARRE)");
        AjouterExemple("HENRI.4 = ROI DE NAVARRE");
        //AjouterExemple("LOUIS_14 = ROI(FRANCE)");
        AjouterExemple("LOUIS.14 = ROI DE FRANCE");
        //AjouterExemple("COULEUR(CHEVAL(ROI(NAVARRE)) ?");
        AjouterExemple("COULEUR DU CHEVAL DU ROI DE NAVARRE ?");
        AjouterExemple("QUELLE EST LA COULEUR DU CHEVAL DU ROI DE NAVARRE ?");
        AjouterExemple("QUELLE EST LA COULEUR DU CHEVAL BLANC DU ROI DE NAVARRE ?");
        AjouterExemple("' Faille : composition incomplète, mauvaise indirection :");
        AjouterExemple("QUELLE EST LA COULEUR DU CHEVAL DU ROI ?");
        AjouterExemple("' Bonne indirection :");
        AjouterExemple("QUEL ROI A UN CHEVAL ?");
        AjouterExemple("' Faille :  mauvaise indirection :");
        AjouterExemple("CHEVAL DU ROI ?");
        AjouterExemple("CHEVAL DU ROI DE NAVARRE ?");
        AjouterExemple("CHEVAL DU ROI DE FRANCE ?");

        AjouterExemple("");
        AjouterExemple("LE CANARI EST UN OISEAU JAUNE");
        AjouterExemple("JAUNE EST UNE COULEUR");
        AjouterExemple("QUEL OISEAU EST JAUNE ?");
        AjouterExemple("QUEL EST L'OISEAU JAUNE ?");
        AjouterExemple("DE QUELLE COULEUR EST LE CANARI ?");
        AjouterExemple("QUELLE EST LA COULEUR DU CANARI ?");
        AjouterExemple("COULEUR CANARI ?");
        AjouterExemple("COULEUR ?");

        AjouterExemple("");
        AjouterExemple("MARSEILLE EST LA VILLE PHOCÉENNE");
        AjouterExemple("DEFERRE EST LE MAIRE DE MARSEILLE");
        AjouterExemple("PHOCÉENNE SIGNIFIE ORIGINAIRE DE PHOCÉE");
        AjouterExemple("LA PHOCÉE EST UNE PROVINCE GRECQUE");
        AjouterExemple("GASTON EST LE PRÉNOM DE DEFERRE");
        AjouterExemple("QUEL EST LE PRÉNOM DU MAIRE DE MARSEILLE ?");
        AjouterExemple("QUEL EST LE PRÉNOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE GRECQUE ?");
        AjouterExemple("' Test du système : composition de fonctions incomplète :");
        AjouterExemple("QUEL EST LE PRÉNOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE ?");
        AjouterExemple("' Autres réponses souhaitées : GASTON EST LE PRÉNOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE PHOCÉENNE");
        AjouterExemple("' ou bien : D'UNE PROVINCE PHOCÉENNE ?");
        AjouterExemple("' ou bien : ** heu... GASTON ?");

        AjouterExemple("");
        AjouterExemple("PAUL POSSÈDE UN PERROQUET BAVARD");
        AjouterExemple("MULTICOLORE SIGNIFIE DE PLUSIEURS COULEURS");
        AjouterExemple("UN PERROQUET EST UN ANIMAL MULTICOLORE");
        AjouterExemple("QUI POSSÈDE UN ANIMAL DE PLUSIEURS COULEURS ?");

        AjouterExemple("");
        AjouterExemple("MARIE EST UNE JOLIE FILLE");
        AjouterExemple("JOLIE EST LE CONTRAIRE DE LAIDE");
        AjouterExemple("MARIE EST-ELLE LAIDE ?");
        AjouterExemple("EST-CE QUE MARIE EST LAIDE ?");

        AjouterExemple("");
        AjouterExemple("SABINE AIME JACQUES");
        AjouterExemple("QUI AIME JACQUES ?");
        AjouterExemple("QUI AIME SABINE ?");
        AjouterExemple("QUI JACQUES AIME-T-IL ?");
        AjouterExemple("' Meilleure réponse : Je L'ignore :");
        AjouterExemple("JACQUES AIME-T-IL SABINE ?");
        AjouterExemple("SABINE AIME-T-ELLE JACQUES ?");

        AjouterExemple("");
        AjouterExemple("' Capacités de la version modifiée :");
        AjouterExemple("' Ordre des termes : DONC TOUT CHEVAL EST HERBIVORE :");
        AjouterExemple("TOUT CHEVAL EST UN ÉQUIDÉ");
        AjouterExemple("OR TOUT ÉQUIDÉ EST HERBIVORE");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("' Sens logique : DONC PAUL EST COMME TOUT HOMME :");
        AjouterExemple("TOUT HOMME EST BIPÈDE");
        AjouterExemple("OR PAUL EST BIPÈDE");
        AjouterExemple("DONC ?");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("' Sens logique : ARISTOTE EST COMME TOUT HOMME :");
        AjouterExemple("TOUT HOMME EST RATIONNEL");
        AjouterExemple("OR ARISTOTE EST RATIONNEL");
        AjouterExemple("DONC ?");
        AjouterExemple("ARISTOTE ÉTAIT GREC");
        AjouterExemple("ARISTOTE ÉTAIT PHILOSOPHE");
        AjouterExemple("PHILOSOPHE GREC ?");

        AjouterExemple("");
        AjouterExemple("' Syllogismes : ordre des assertions");
        AjouterExemple("OR SOCRATE EST UN HOMME");
        AjouterExemple("TOUT HOMME EST MORTEL");
        AjouterExemple("DONC ?");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("' Syllogismes : plusieurs à la fois");
        AjouterExemple("TOUT HOMME EST MORTEL");
        AjouterExemple("TOUT HOMME EST BIPÈDE");
        AjouterExemple("OR SOCRATE EST UN HOMME");
        AjouterExemple("DONC ?");
        AjouterExemple("OR PAUL EST UN HOMME");
        AjouterExemple("DONC ?");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("TOUT LOGICIEN EST INCOMPRIS");
        AjouterExemple("OR TOUT HOMME SENSÉ EST LOGICIEN");
        AjouterExemple("DONC ?");

        AjouterExemple("");
        AjouterExemple("' Syllogismes : syntaxe imparfaite (de la version modifiée)");
        AjouterExemple("LE LOUVRE EST BEAU");
        AjouterExemple("OR J'AIME TOUT CE QUI EST BEAU");
        AjouterExemple("DONC ?");
        AjouterExemple("");

        }
    }
}

