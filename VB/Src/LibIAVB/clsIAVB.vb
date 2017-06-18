
' Fichier clsIAVB.vb : Classe du moteur IAVB : Intelligence Artificielle en Visual Basic
' ------------------

' Documentation : IAVB.html
' http://patrice.dargenton.free.fr/ia/iavb/IAVB.html
' http://patrice.dargenton.free.fr/ia/iavb/IAVB.vbproj.html
' https://github.com/PatriceDargenton/IAVB
' Version 3.12 du 18/06/2017
' Version 3.10 du 06/12/2009
' Version 3.01 du 02/09/2007

' IAVB est la transcription en Visual Basic d'un logiciel paru dans la revue
'  MICRO-SYSTEMES en Décembre 1984, pages 195-202 :
'  "Mini-système expert pour Apple II" par Philippe LARVET :
'  Gestion d'une base de connaissances.

' Par Patrice Dargenton : mailto:patrice.dargenton@free.fr
' http://patrice.dargenton.free.fr/index.html
' http://patrice.dargenton.free.fr/CodesSources/index.html

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' u pour Unsigned (non signé : entier positif)
' a pour Array (tableau) : ()
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------

Imports System.Text ' Pour StringBuilder
Imports System.Collections.Generic

Namespace IAVB

Public Class clsIAVB

#Region "Interface"

' Booléen pour forcer le fonctionnement dans la version originale (1984)
'  ou bien modifiée (2001)
Public m_bVersionModifiee As Boolean = True

' Normaliser les espaces et sauts de ligne en sortie
Public m_bNormalisationSortieTrimEtVbLf As Boolean = True

' Convertir tout en minuscules (pour éviter d'avoir à taper des majuscules avec accent,
'  et pour pouvoir prononcer "ça" via MS-Agent)
Public m_bTraiterEnMinuscules As Boolean = True

' 29/04/2017 Enlever les accents (pour pouvoir comparer avec la version originale de 1984)
Public m_bTraiterSansAccents As Boolean = True

#End Region

#Region "Constantes"

Public ReadOnly vbLf$ = VBChr(10) ' Chr(10) ' CSharp : mettre "\n"

Private Const sTitreMsg$ = "Moteur IAVB3"

' iNbMotsBCMax : Nbre de mots signifiants susceptibles
'  d'être conservé en mémoire pour chaque assertion (au moins 3).
'  seul les 2 premiers mots peuvent former une clé pour une relation H ou une indirection
Private Const iNbMotsBCMax% = 4

' Nbre de mots signifiants de l'entrée max.
' Peut être > iNbMotsBCMax notamment pour les compositions de fonctions
Private Const iNbMotsEntreeMax% = 10

Private Const sEspace$ = " "

Private sMotQU, sMotDonc, sMotEstComme, sMotEst, sMotSont As String
Private sMotTou, sMotUn, sMotOr, sMotQuelque, sMotQuelques As String

Private Const sGM$ = """"
Private Const sPointInterrog$ = "?"
Private Const sPointExclam$ = "!"
Private Const sSlash$ = "/"
Private Const sQuote$ = "'"
Private Const sTiret$ = "-"
Private Const sEgal$ = "="
Private Const sParenthOuv$ = "("
Private Const sParenthFerm$ = ")"
Private Const sDeuxPoints$ = ":"

Private Const sCmdSilenceOk$ = "Okay, je me tais."
Private Const sCmdParlerOk$ = "Je vous écoute."
Private Const sCmdCopieOk$ = "La discusion a été copiée dans le presse papier."
Private Const sCmdCopiePb$ = "La discusion n'a pas pu être copiée dans le presse papier !"
Private Const sCmdEffacerConfirm$ = "Voulez-vous effacer toute la base ?"
Private Const sCmdEffacerOk$ = "Base entièrement effacée !"
Private Const sCmdSupprAssertionOk1$ = "Assertion " ' & I & " supprimée"
Private Const sCmdSupprAssertionOk2$ = " supprimée."
Private Const sCmdSupprAssertionPb$ = "Numéro de l'assertion hors limites !"
Private Const sReponseCompris$ = "Compris." '"Compris"
Private Const sReponseConnaisPas$ = "Je ne connais pas "
Private Const sReponseOui$ = "Oui." '"Oui"
Private Const sReponseRienConclure$ = "Je ne peux rien conclure !"
Private Const sReponseIgnore$ = "Je l'ignore." '"Je l'ignore"
Private Const sReponseNon$ = "Non." '"Non"
Private Const sReponseRechApprof$ = "Recherche approfondie..."
Private Const sErrAssertionDejaConnue$ = "Assertion déjà connue !"
'Private Const sErrBasePleine$ = "Stop ! La base est pleine !"
Private Const sErrPhraseTropCourte$ = "Votre phrase est trop courte !"
Private Const sErrPhraseTropLongue$ = "Votre phrase est trop longue !"
Private Const sErrPhraseIncomplete$ = "Votre phrase est incomplète !"
Private Const sErrCmdInconnue$ = "Commande inconnue !"
Private Const sErrBaseVide$ = "Base vide !"

#End Region

#Region "Classes"

Private Class TAssertion
    Public iNbMots%
    Public asMot$()
    Public iPosFinMot1%
    Public iPosFinMot2%
    Public sEntree$
End Class

Private Class TParamEntree

    Public bMotQuestion As Boolean
    Public bInterrogation As Boolean
    Public bFinTraitement As Boolean
    Public sEntree$
    Public sEntreeCompilee$ ' Concaténation de tous les asMotsEntree$(i)
    Public iNbMotsEntree% ' Nbre de mots signifiants extraits de sEntree
    Public asMotsEntree$() ' Mots signifiants extraits de l'assertion
    Public iPosFinMot1%, iPosFinMot2%

    ' n° de l'assertion contenant le mot signifiant n°K de l'entrée,
    '  tableau des K mots de l'entrée
    '  D'abord la première assertion : ControleExistenceMots(),
    '  puis les assertions suivantes : bAssertCMot() pour les indirections.
    Public aiNumAssertCMot%()

    ' Conservation de l'indice minimum pour
    '  l'espace de recherche max. de l'assertion
    '  (contenant un des mots de l'assertion en cours)
    Public iNumAssertMinRech%

End Class

#End Region

#Region "Déclarations"

' Liste des mots ignorés en général (non-signifiants)
Dim m_lstMotsIgnores As New List(Of String) ' Dim = Private

Dim m_lstBC As New List(Of TAssertion) ' Base de Connaissances (KB)

Dim m_iNumAssertionEnCours%

Dim m_bQuestionPreced_Donc As Boolean

' Préparer les réponses multiples pour les renvoyer en une seule chaine de caractères
Private m_lsReponses As List(Of String)
' Pour pouvoir copier la discussion dans le presse papier
Private m_lsReponsesTot As New List(Of String)
' Pour pouvoir synchroniser les réponses vocales avec les réponses écrites
Private m_lsReponsesVocales As List(Of String)

#End Region

#Region "Interface"

Public m_sRappelQuestion$ = ""
Public m_sReponse$ = ""
Public m_bCopierPressePapier As Boolean = False
Public m_sDiscussion$ = ""

Public sCmdEffBase, sCmdEff, sCmdLister As String
Public sCmdCopier, sCmdSilence, sCmdParler As String

Public m_bVoixActive As Boolean
Public m_sReponseVocale$ = ""

#End Region

#Region "Fonctions publiques"

Public Sub Initialiser(bVersionModifiee As Boolean, bNormalisationSortieTrimEtVbLf As Boolean, _
    bTraiterEnMinuscules As Boolean, bTraiterSansAccents As Boolean)

    m_bVersionModifiee = bVersionModifiee
    m_bNormalisationSortieTrimEtVbLf = bNormalisationSortieTrimEtVbLf
    m_bTraiterEnMinuscules = bTraiterEnMinuscules
    m_bTraiterSansAccents = bTraiterSansAccents

    m_lstMotsIgnores = New List(Of String)

    m_lstBC = New List(Of TAssertion)

    InitCmd()
    InitMotsIgnores()

End Sub

Private Sub AjouterAssertion()
    Dim assert As New TAssertion
    ReDim assert.asMot(0 To iNbMotsBCMax)
    For i As Integer = 0 To iNbMotsBCMax
        assert.asMot(i) = ""
    Next
    m_lstBC.Add(assert)
End Sub

Public Sub IAVBMain(sEntree$)

    m_sReponseVocale = ""
    m_sDiscussion = ""
    m_bCopierPressePapier = False
    m_sRappelQuestion = "" : m_sReponse = ""
    m_lsReponses = New List(Of String)
    m_lsReponsesVocales = New List(Of String)

    Dim prm As New TParamEntree
    prm.sEntree = sEntree
    AnalyseEntree(prm)
    If prm.bFinTraitement Then Exit Sub

    Dim bRechecheApprofondie As Boolean
    If prm.bInterrogation And prm.asMotsEntree(1) = sMotDonc Then
        bRechecheApprofondie = False
        If m_bVersionModifiee And m_bQuestionPreced_Donc Then
            bRechecheApprofondie = True
            If m_iNumAssertionEnCours > 0 Then ReponseIAVB( _
                m_lstBC(m_iNumAssertionEnCours).sEntree & sEspace & sDeuxPoints, bListe:=True)
            ReponseIAVB(sReponseRechApprof, bAffiche2Etoiles:=True, bListe:=True)
        End If
        Syllogisme(bRechecheApprofondie)
        m_bQuestionPreced_Donc = True
        Exit Sub
    End If
    m_bQuestionPreced_Donc = False

    If Not prm.bInterrogation Then
        If prm.iNbMotsEntree <= 1 Then
            ReponseIAVB(sErrPhraseTropCourte, bAffiche2Etoiles:=True)
            Exit Sub
        End If
        AjoutBase(prm) : Exit Sub
    End If

    ' Interrogation : controle existence des mots
    ControleExistenceMots(prm)
    If prm.bFinTraitement Then Exit Sub

    ' A vérifier : si prm.iNbMotsEntree > iNbMotsBCMax : relationH après réduction ?
    If prm.iNbMotsEntree <= iNbMotsBCMax Then
        If bRelationHorizontale(prm) Then Exit Sub
    End If

    ' Composition de fcts (relation verticale)
    If bComposFonction(prm) Then Exit Sub

    If bIndirection(prm) Then Exit Sub

    ' Echec Final
    If prm.bMotQuestion Then
        ReponseIAVB(sReponseIgnore, bAffiche2Etoiles:=True)
        Exit Sub
    End If
    ReponseIAVB(sReponseNon, bAffiche2Etoiles:=True)

End Sub

#End Region

#Region "Moteur IAVB"

Private Sub ReponseIAVB(ByRef sTexte$, _
    Optional bAffiche2Etoiles As Boolean = False, _
    Optional bRappelReponse As Boolean = False, _
    Optional bSilence As Boolean = False, _
    Optional sTexteParleSpecifique$ = "", _
    Optional bListe As Boolean = False, _
    Optional bFinListe As Boolean = False)

    If bFinListe Then
        ' Remplacer liste de String par une String & vbLf
        Dim sb As New StringBuilder
        For Each sReponse As String In m_lsReponses
            ' AppendLine ne convient pas
            If m_bNormalisationSortieTrimEtVbLf Then
                sb.Append(sReponse.Trim & vbLf) ' 09/04/2017 Normalisation des réponses
            Else
                sb.Append(sReponse & vbLf)
            End If
        Next
        m_sReponse = sb.ToString

        sb = New StringBuilder
        For Each sReponse As String In m_lsReponsesVocales
            sb.Append(sReponse & vbLf)
        Next
        m_sReponseVocale = sb.ToString

        Exit Sub
    End If

    Dim sTexteEcrit$ = sTexte
    If bAffiche2Etoiles Then sTexteEcrit = "** " & sTexte

    ' Affichage du texte avant de parler
    If bRappelReponse Then m_sRappelQuestion = sTexteEcrit

    Dim sTexteParle$
    If m_bVoixActive And Not bSilence Then
        sTexteParle = sTexte
        If sTexteParleSpecifique.Length > 0 Then sTexteParle = sTexteParleSpecifique
        ' Il peut y avoir une liste de réponses vocales
        '  pb : on ne connait pas la fin de la liste !
        '  Soluce : utiliser la même fin de liste qu'a l'affichage,
        '   et fait paire de réponses parlée/écrite : synchro + facile
        m_lsReponsesVocales.Add(sTexteParle)
        m_lsReponsesVocales.Add(sTexteEcrit) ' Pour la synchro.

        ' Préparer tjrs la réponse finale
        Dim sb As New StringBuilder
        For Each sReponse As String In m_lsReponsesVocales
            sb.Append(sReponse & vbLf)
        Next
        m_sReponseVocale = sb.ToString

    End If

    ' 09/04/2017 Normalisation des réponses
    If m_bNormalisationSortieTrimEtVbLf Then sTexteEcrit &= vbLf

    If Not bRappelReponse Then

        ' bSilence : commentaire, cmd, ou ligne vide : là aussi on se contente
        '  de répéter l'entrée : comme rappel question
        If bSilence Then
            m_sRappelQuestion = sTexteEcrit
        ElseIf bListe Then
            m_lsReponses.Add(sTexteEcrit)
        Else
            m_sReponse = sTexteEcrit
        End If

    End If

    m_lsReponsesTot.Add(sTexteEcrit)

End Sub

Private Sub AnalyseEntree(prm As TParamEntree)

    prm.bInterrogation = False
    prm.bInterrogation = False
    prm.bFinTraitement = False
    prm.bMotQuestion = False
    prm.iNbMotsEntree = 0
    prm.sEntreeCompilee = ""

    ' Extraction des mots
    ReDim prm.asMotsEntree(0 To iNbMotsEntreeMax)
    ReDim prm.aiNumAssertCMot(0 To iNbMotsEntreeMax)
    For K As Integer = 1 To iNbMotsBCMax
        prm.asMotsEntree(K) = "" : prm.aiNumAssertCMot(K) = 0
    Next K

    Dim iLenEntree% = iVBLen(prm.sEntree)

    ' Examen de l'entrée
    If iLenEntree = 0 Then ' Ligne vide
        ReponseIAVB(prm.sEntree, bSilence:=True)
        prm.bFinTraitement = True : Exit Sub
    End If
    ' Gestion des commentaires
    If sVBLeft(prm.sEntree, 1) = sQuote Then
        ReponseIAVB(prm.sEntree, bSilence:=True)
        prm.bFinTraitement = True : Exit Sub
    End If

    ' 29/04/2017 Enlever les accents, pour pouvoir comparer avec la version originale
    If m_bTraiterSansAccents Then prm.sEntree = sEnleverAccents(prm.sEntree)

    ' Convertir tout en minuscules pour pouvoir prononcer "ça"
    If m_bTraiterEnMinuscules Then prm.sEntree = prm.sEntree.ToLowerInvariant()

    If sVBLeft(prm.sEntree, 1) = sSlash Then
        ReponseIAVB(prm.sEntree, bSilence:=True)
        TraiterCmd(prm, iLenEntree)
        Exit Sub
    End If

    ' Afficher immédiatement le texte, sans attendre qu'il soit prononcé
    ReponseIAVB(prm.sEntree, bRappelReponse:=True)

    If sVBRight(prm.sEntree, 1) = sPointInterrog Then
        prm.sEntree = sVBLeft(prm.sEntree, iLenEntree - 1)
        iLenEntree = iVBLen(prm.sEntree)
        prm.bInterrogation = True
    End If

    AnalyseMots(prm, iLenEntree)

End Sub

Private Sub AnalyseMots(prm As TParamEntree, iLenEntree%)

    Dim iPos% = 1 ' Boucle sur les lettres
    Do
        Dim iMemPos% = iPos
        Dim sLettreEntree$ = ""

        Do
            iPos += 1
            sLettreEntree = sVBMid(prm.sEntree, iPos, 1)
        Loop While iPos <= iLenEntree AndAlso sLettreEntree <> sEspace AndAlso _
            sLettreEntree <> sQuote AndAlso sLettreEntree <> sTiret AndAlso _
            sLettreEntree <> sParenthOuv AndAlso sLettreEntree <> sParenthFerm

        ' Chaque mot extrait de sEntree
        Dim sMotEntree$ = sVBMid(prm.sEntree, iMemPos, iPos - iMemPos)
        If iMemPos = 1 AndAlso sVBLeft(sMotEntree, 2) = sMotQU Then _
            prm.bMotQuestion = True : GoTo RechercheMotSuivant

        Dim bMotIgnore As Boolean = False
        For Each sMotIgnore As String In m_lstMotsIgnores
            If sMotIgnore = sMotEntree Then bMotIgnore = True : Exit For
        Next
        If bMotIgnore Then GoTo RechercheMotSuivant

        ' Ajout d'un mot signifiant
        prm.iNbMotsEntree += 1
        ' Modification : test aussi avec iNbMotsEntreeMax
        If (Not prm.bInterrogation AndAlso prm.iNbMotsEntree > iNbMotsBCMax) OrElse _
            prm.iNbMotsEntree > iNbMotsEntreeMax Then
            ReponseIAVB(sErrPhraseTropLongue, bAffiche2Etoiles:=True)
            prm.bFinTraitement = True : Exit Sub
        End If

        prm.sEntreeCompilee = prm.sEntreeCompilee & sMotEntree
        prm.asMotsEntree(prm.iNbMotsEntree) = sMotEntree
        If prm.iNbMotsEntree = 1 Then prm.iPosFinMot1 = iPos
        If prm.iNbMotsEntree = 2 Then prm.iPosFinMot2 = iPos

RechercheMotSuivant:
        If iPos > iLenEntree Then Exit Sub

        Do
            iPos += 1
        Loop While sVBMid(prm.sEntree, iPos, 1) = sEspace AndAlso iPos <= iLenEntree

    Loop While iPos <= iLenEntree

    If prm.asMotsEntree(1) = "" Then
        ReponseIAVB(sErrPhraseIncomplete, bAffiche2Etoiles:=True)
        prm.bFinTraitement = True
    End If

End Sub

Private Sub TraiterCmd(prm As TParamEntree, iLenEntree%)

    If sVBLeft(prm.sEntree, sCmdLister.Length) = sCmdLister Then ListerBase(prm) : Exit Sub
    If sVBLeft(prm.sEntree, sCmdEffBase.Length) = sCmdEffBase Then EffacerBase(prm) : Exit Sub
    If sVBLeft(prm.sEntree, sCmdEff.Length) = sCmdEff Then SupprimerAssertion(prm, iLenEntree) : Exit Sub
    If sVBLeft(prm.sEntree, sCmdCopier.Length) = sCmdCopier Then CopierPressePapier(prm) : Exit Sub

    If sVBLeft(prm.sEntree, sCmdSilence.Length) = sCmdSilence Then
        ReponseIAVB(sCmdSilenceOk, bAffiche2Etoiles:=True)
        m_bVoixActive = False
        prm.bFinTraitement = True : Exit Sub
    End If

    If sVBLeft(prm.sEntree, sCmdParler.Length) = sCmdParler Then
        m_bVoixActive = True
        ReponseIAVB(sCmdParlerOk, bAffiche2Etoiles:=True)
        prm.bFinTraitement = True : Exit Sub
    End If

    ReponseIAVB(sErrCmdInconnue, bAffiche2Etoiles:=True)
    prm.bFinTraitement = True

End Sub

Private Sub ListerBase(prm As TParamEntree)

    Dim iNbAssertions% = m_lstBC.Count
    If iNbAssertions = 0 Then
        ReponseIAVB(sErrBaseVide, bAffiche2Etoiles:=True) ', bListe:=True)
        'GoTo Fin
        prm.bFinTraitement = True
        Exit Sub
    End If

    For I As Integer = 0 To iNbAssertions - 1
        ReponseIAVB((I + 1) & sEspace & m_lstBC(I).sEntree, bListe:=True)
    Next I

Fin:
    ReponseIAVB("", bFinListe:=True)
    prm.bFinTraitement = True

End Sub

Private Sub CopierPressePapier(prm As TParamEntree)

    ' Copie du résultat d'exécution dans le presse papier
    Dim sbDisc As New StringBuilder()
    For Each sReponse As String In m_lsReponsesTot
        sbDisc.AppendLine(sReponse)
    Next
    m_sDiscussion = sbDisc.ToString

    ' La copie dans le presse-papier doit maintenant être faite depuis l'appelant
    '  on signale simplement ici que l'on a compris la commande
    m_bCopierPressePapier = True
    'ReponseIAVB(sCmdCopieOk, bAffiche2Etoiles:=True)
    prm.bFinTraitement = True

End Sub

Public Sub CopiePressePapierOk()
    ReponseIAVB(sCmdCopieOk, bAffiche2Etoiles:=True)
End Sub

Public Sub CopiePressePapierEchec()
    ReponseIAVB(sCmdCopiePb, bAffiche2Etoiles:=True)
End Sub

Private Sub SupprimerAssertion(prm As TParamEntree, iLenEntree%)

    Dim i% = m_lstBC.Count
    If i = 0 Then
        ReponseIAVB(sErrBaseVide, bAffiche2Etoiles:=True)
        prm.bFinTraitement = True
        Exit Sub
    End If
    i -= 1 ' Par défaut effacer la dernière assertion

    If sVBMid(prm.sEntree, 3, 1) <> sEspace Then
        Dim sNumAssert$ = sVBRight(prm.sEntree, iLenEntree - 2)
        Dim iPosCom% = sNumAssert.IndexOf("'")
        If iPosCom > -1 Then sNumAssert = sNumAssert.Substring(0, iPosCom)
        ' ' Par défaut effacer la dernière assertion
        If sNumAssert.Trim = "" Then GoTo Suite
        i = iVBVal(sNumAssert)
        i -= 1 ' Indice 0 maintenant
        If i < 0 Or i >= m_lstBC.Count Then
            Dim sReponse$ = sCmdSupprAssertionPb & " " & i + 1 & " <> (" & 1 & ", " & m_lstBC.Count & ")"
            ReponseIAVB(sReponse, bAffiche2Etoiles:=True)
            prm.bFinTraitement = True
            Exit Sub
        End If
    End If

Suite:
    m_lstBC.RemoveAt(i)

    ReponseIAVB(sCmdSupprAssertionOk1 & (i + 1) & sCmdSupprAssertionOk2, bAffiche2Etoiles:=True)

    prm.bFinTraitement = True

End Sub

Private Sub EffacerBase(prm As TParamEntree)

    Dim iNbAssertions% = m_lstBC.Count
    If iNbAssertions = 0 Then
        ReponseIAVB(sErrBaseVide, bAffiche2Etoiles:=True)
        prm.bFinTraitement = True : Exit Sub
    End If
    ' Pb : pas de confirmation vocale
    'ReponseIAVB(sCmdEffacerConfirm, bAffiche2Etoiles:=True)

    If Not bQuestionResponseOui(sCmdEffacerConfirm) Then _
        prm.bFinTraitement = True : Exit Sub

    m_lstBC = New List(Of TAssertion)
    ReponseIAVB(sCmdEffacerOk, bAffiche2Etoiles:=True)
    iNbAssertions = 0
    prm.bFinTraitement = True

End Sub

Private Sub AjoutBase(prm As TParamEntree)

    ' Assertion : Controle existence de l'assertion dans la base
    Dim iNbAssertions% = m_lstBC.Count
    For I As Integer = 0 To iNbAssertions - 1
        Dim sGdTerme$ = ""
        For J As Integer = 1 To iNbMotsBCMax
            sGdTerme &= m_lstBC(I).asMot(J)
        Next J
        If sGdTerme = prm.sEntreeCompilee Then
            If m_bVersionModifiee Then m_iNumAssertionEnCours = I
            ReponseIAVB(sErrAssertionDejaConnue, bAffiche2Etoiles:=True)
            Exit Sub
        End If
    Next I

    AjouterAssertion()
    Dim I1% = m_lstBC.Count - 1
    For J As Integer = 1 To iNbMotsBCMax
        m_lstBC(I1).asMot(J) = prm.asMotsEntree(J)
    Next J
    Dim assert As TAssertion = m_lstBC(I1)
    assert.sEntree = prm.sEntree
    assert.iNbMots = prm.iNbMotsEntree
    assert.iPosFinMot1 = prm.iPosFinMot1
    assert.iPosFinMot2 = prm.iPosFinMot2
    ReponseIAVB(sReponseCompris, bAffiche2Etoiles:=True)
    m_iNumAssertionEnCours = m_lstBC.Count - 1

End Sub

Private Sub ControleExistenceMots(prm As TParamEntree)

    ' Interrogation : contrôle de l'existence des mots

    Dim iNbAssertions% = m_lstBC.Count
    Dim iMemNumAssertionEnCours% = 0
    m_iNumAssertionEnCours = iNbAssertions - 1
    Dim iLenEntree% = iVBLen(prm.sEntree)
    prm.iNumAssertMinRech = iNbAssertions - 1 ' Utile pour la version modifiée

    For K As Integer = 1 To prm.iNbMotsEntree

        Dim bMotPresentAssertion As Boolean = False
        For I As Integer = 0 To iNbAssertions - 1
            Dim bMotTrouve As Boolean = False
            Dim J% = 1
            For J = 1 To m_lstBC(I).iNbMots
                If m_lstBC(I).asMot(J) = prm.asMotsEntree(K) Then bMotTrouve = True : Exit For
            Next J
            If Not bMotTrouve Then Continue For

            bMotPresentAssertion = True
            If prm.aiNumAssertCMot(K) = 0 Then prm.aiNumAssertCMot(K) = I

            If prm.iNbMotsEntree > 1 Then iMemNumAssertionEnCours = I : Exit For

            If J = 2 And iLenEntree > iVBLen(prm.asMotsEntree(K)) + 1 Then
                ReponseIAVB(sVBLeft(m_lstBC(I).sEntree, m_lstBC(I).iPosFinMot1), bListe:=True)
                ' Question spécifique, ex.:
                ' ET LAQUELLE EST FOFOLLE ?
                ' MINNA = Left(aBC(I).sEntree, aBC(I).iPosFinMot1)
                ' aBC(I).sEntree = MINNA EST FOFOLLE
                Continue For 'GoTo AssertionSuivante
            End If
            ReponseIAVB(m_lstBC(I).sEntree, bListe:=True)
            ' Entrée contenant le mot K : Question générique, ex.:
            ' RESPONSABLE ?
            ' M.BERTRAND EST RESPONSABLE DE L'ANNEXE
            ' M.JACQUES EST RESPONSABLE DE LA SAISIE
        Next I

        If Not bMotPresentAssertion Then
            Dim sReponse$ = sReponseConnaisPas & prm.asMotsEntree(K)
            ReponseIAVB(sReponse, bAffiche2Etoiles:=True)
            prm.bFinTraitement = True
            Exit Sub ' Pas de liste ici, quitter directement
        End If
        If prm.iNbMotsEntree = 1 Then
            prm.bFinTraitement = True
            GoTo Fin ' Liste possible ici
        End If

        ' Conserver l'indice minimum pour l'espace de recherche max. de l'assertion
        If m_bVersionModifiee Then
            If iMemNumAssertionEnCours < prm.iNumAssertMinRech Then _
                prm.iNumAssertMinRech = iMemNumAssertionEnCours
        Else
            If iMemNumAssertionEnCours < m_iNumAssertionEnCours Then _
                m_iNumAssertionEnCours = iMemNumAssertionEnCours
        End If

    Next K

Fin:
    ReponseIAVB("", bFinListe:=True)

End Sub

Private Function bRelationHorizontale(prm As TParamEntree) As Boolean

    Dim aiMotsTrouves%(2)
    Dim bRelationHorizontale0 As Boolean = False
    Dim iDeb% = m_iNumAssertionEnCours
    If m_bVersionModifiee Then iDeb = prm.iNumAssertMinRech

    Dim iNbAssertions% = m_lstBC.Count
    For I As Integer = iDeb To iNbAssertions - 1
        Dim iNbMotsTrouves% = 0
        Dim iMemJ% = 1

        For K As Integer = 1 To prm.iNbMotsEntree
            Dim bMotTrouve As Boolean = False
            Dim J%
            For J = iMemJ To m_lstBC(I).iNbMots
                If m_lstBC(I).asMot(J) = prm.asMotsEntree(K) Then bMotTrouve = True : Exit For
            Next J
            If bMotTrouve Then
                iNbMotsTrouves += 1
                If iNbMotsTrouves <= 2 Then aiMotsTrouves(iNbMotsTrouves) = J
            End If
            iMemJ = J ' Force l'ordre de recherche à Sujet/Verbe/Complement seulement
            ' Cela permet de répondre Oui à la question :
            ' SABINE AIME-T-ELLE JACQUES ?
            ' et cela empêche de trouver une relation H à la question :
            ' JACQUES AIME-T-IL SABINE ?
        Next K

        If iNbMotsTrouves = 0 Then GoTo AssertionSuivante

        If iNbMotsTrouves = prm.iNbMotsEntree And iNbMotsTrouves = m_lstBC(I).iNbMots Then
            ReponseIAVB(sReponseOui, bAffiche2Etoiles:=True)
            Return True
        End If

        If iNbMotsTrouves <> prm.iNbMotsEntree Then GoTo AssertionSuivante
        bRelationHorizontale0 = True

        If aiMotsTrouves(1) = 1 And aiMotsTrouves(2) = 2 Then
            Dim iLenE% = iVBLen(m_lstBC(I).sEntree)
            ReponseIAVB(sVBRight(m_lstBC(I).sEntree, iLenE - m_lstBC(I).iPosFinMot2), bListe:=True)
            ' Ex.:
            ' QUI MARIE REGARDE-T-ELLE ?
            ' Réponse : HUGUES
            ' aBC(I).sEntree = MARIE REGARDE HUGUES
            GoTo AssertionSuivante
        End If

        If aiMotsTrouves(1) = 2 Then
            ReponseIAVB(sVBLeft(m_lstBC(I).sEntree, m_lstBC(I).iPosFinMot1), bListe:=True)
            ' Ex.:
            ' CHEF DE SERVICE ?
            ' M.RENE
            ' M.DUBOIS
            ' aBC(I).sEntree = M.RENE EST LE CHEF DU SERVICE COMPTABILITE
            ' aBC(I).sEntree = M.DUBOIS EST CHEF DU SERVICE PHOTO
            GoTo AssertionSuivante
        End If

        If aiMotsTrouves(1) = 3 Then
            ReponseIAVB(sVBLeft(m_lstBC(I).sEntree, m_lstBC(I).iPosFinMot2), bListe:=True)
            ' Ex.:
            ' SERVICE PHOTO ?
            ' M.DUBOIS EST CHEF
            GoTo AssertionSuivante
        End If

        ReponseIAVB(m_lstBC(I).sEntree, bListe:=True) ' Liste
        ' Pas d'exemple !

AssertionSuivante:
    Next I

    ReponseIAVB("", bFinListe:=True)

    Return bRelationHorizontale0

End Function

Private Function bComposFonction(prm As TParamEntree) As Boolean

    ' Composition de fonction (relation verticale)

    Dim asMotReduit$(iNbMotsEntreeMax + 1)
    For K As Integer = 1 To prm.iNbMotsEntree
        asMotReduit(K) = prm.asMotsEntree(K)
    Next K

    Dim iNbAssertions% = m_lstBC.Count
    For K As Integer = prm.iNbMotsEntree To 2 Step -1
        Dim sGdTerme$ = asMotReduit(K - 1) & asMotReduit(K)
        Dim bCompFct As Boolean = False

        ' Recherche des gd termes sur les mots n° 2 et 3 seulement de la BC :
        '  du type Y = F(X) : Y:1, F:2 et X:3, par ex.: BLANC = COULEUR(ARTABAN)
        Dim iDeb% = m_iNumAssertionEnCours
        If m_bVersionModifiee Then iDeb = prm.iNumAssertMinRech
        Dim I%
        For I = iDeb To iNbAssertions - 1
            If m_lstBC(I).asMot(2) & m_lstBC(I).asMot(3) = sGdTerme Then bCompFct = True : Exit For
        Next I
        If Not bCompFct Then Return False

        asMotReduit(K - 1) = m_lstBC(I).asMot(1)
    Next K

    ' Fin de composition de fct : réponse trouvée
    ReponseIAVB(asMotReduit(1))
    Return True
    ' Ex.:
    ' ARTABAN = CHEVAL(HENRI_IV)
    ' BLANC = COULEUR(ARTABAN)
    ' HENRI_IV = ROI(NAVARRE)
    ' COULEUR(CHEVAL(ROI(NAVARRE))) ?
    ' Réponse : BLANC
    ' sGdTerme = COULEURARTABAN

End Function

Private Function bIndirection(prm As TParamEntree) As Boolean

    Dim iMemNumAssertCMot2% = prm.aiNumAssertCMot(2)
    Dim bIndirection0 As Boolean = False

    Do
        prm.aiNumAssertCMot(2) = iMemNumAssertCMot2
        Do
            Dim I1% = prm.aiNumAssertCMot(1)
            Dim I2% = prm.aiNumAssertCMot(2)
            For J1 As Integer = 1 To m_lstBC(I1).iNbMots
                Dim sGdTerme$ = m_lstBC(I1).asMot(J1)
                For J2 As Integer = 1 To m_lstBC(I2).iNbMots
                    If sGdTerme <> m_lstBC(I2).asMot(J2) Then GoTo MotSuivant
                    If I1 <> I2 And sGdTerme <> prm.asMotsEntree(1) And _
                        sGdTerme <> prm.asMotsEntree(2) And _
                        (prm.iNbMotsEntree <= 3 Or Not m_bVersionModifiee) Then
                        ' Modification : il ne peut y avoir indirection que sur une
                        '  question simple (au plus 3 termes signifiants)
                        bIndirection0 = True
                        ReponseIAVB(m_lstBC(I1).asMot(J1), bListe:=True)
                        J1 = iNbMotsBCMax : J2 = iNbMotsBCMax ' Exit For 2x

                        ' Ex.: QUELLE EST LA COULEUR DU STYLO DE FRANÇOIS ?
                        ' Réponse : BLEU
                        ' sGdTerme = BLEU, asMotsEntree(1) = COULEUR, asMotsEntree(2) = STYLO

                        ' Ex.: QUELLE FILLE EST SAGE ?
                        ' Réponse : ANNIE
                        ' sGdTerme = ANNIE, asMotsEntree(1) = FILLE, asMotsEntree(2) = SAGE

                        ' ARTABAN = CHEVAL(HENRI_IV)
                        ' BLANC = COULEUR(ARTABAN)
                        ' HENRI_IV = ROI(NAVARRE)
                        ' Bonne indirection:
                        ' QUEL ROI A UN CHEVAL ? : HENRI_IV
                        ' Faille :  mauvaise indirection
                        ' CHEVAL DU ROI ? : HENRI_IV
                        ' QUELLE EST LA COULEUR DU CHEVAL DU ROI ? : ARTABAN
                        ' sGdTerme = ARTABAN
                        ' prm.iNbMotsEntree = 3 : COULEUR, CHEVAL, ROI
                        ' m_aBC(I1).asMot(J1) = ARTABAN
                        ' prm.asMotsEntree(1) = COULEUR
                        ' prm.asMotsEntree(2) = CHEVAL

                    End If
MotSuivant:
                Next J2
            Next J1
        Loop While bAssertCMot(prm, K:=2)
    Loop While bAssertCMot(prm, K:=1)

    ReponseIAVB("", bFinListe:=True)

    Return bIndirection0

End Function

Private Function bAssertCMot(prm As TParamEntree, K%) As Boolean

    ' Module de recherche d'un prochain aiNumAssertCMot(K)
    Dim iDeb% = prm.aiNumAssertCMot(K)
    Dim iNbAssertions% = m_lstBC.Count
    For I As Integer = iDeb + 1 To iNbAssertions - 1
        For J As Integer = 1 To m_lstBC(iDeb).iNbMots
            If m_lstBC(I).asMot(J) = prm.asMotsEntree(K) Then
                prm.aiNumAssertCMot(K) = I
                ' Ex.:ANIMAL MANGEUR ?
                ' Réponse : CHAT, TIGRE
                ' asMotsEntree(K) = MANGEUR dans les 2 cas
                Return True
            End If
        Next J
    Next I
    Return False

End Function

Private Sub Syllogisme(bRechecheApprofondie As Boolean)

    ' Recherche du moyen terme
    ' sPetitTerme$ : Petit terme extrait de l'assertion pour syllogisme
    ' sGdTerme$ : Concaténation de 2 ou 3 mots signifiants et extrait de l'assertion

    Dim bPasse2 As Boolean = False

    Dim iNbAssertions% = m_lstBC.Count
    Dim iNumAss2% = iNbAssertions - 1
    If m_bVersionModifiee Then iNumAss2 = m_iNumAssertionEnCours ' Màj si assertion déjà connue

    Dim bAucunMotTrouve As Boolean = True
    Dim iAssertionPreced% = iNumAss2 - 1
    ' Si bVersionModifiee seulement :
    If bRechecheApprofondie Then iAssertionPreced = 0 ' Recherche approfondie
    Dim iPas% = -1
    Dim iDebut% = iNumAss2 - 1
    Dim iFin% = iAssertionPreced

Passe2:
    If bPasse2 Then
        iPas = 1
        iDebut = iNumAss2 + 1
        iFin = iNbAssertions - 1
    End If

    For iNumAss1 As Integer = iDebut To iFin Step iPas
        Dim bMotTrouve As Boolean = False
        Dim iLenA1% = iVBLen(m_lstBC(iNumAss1).sEntree)
        Dim iLenA2% = iVBLen(m_lstBC(iNumAss2).sEntree)
        Dim iPosMT_PMaj% = 0 ' Place du moyen terme dans la prémisse majeure
        Dim iPosMT_PMin% = 0 ' Place du moyen terme dans la prémisse mineure
        Dim sMotPivot$ = ""
        For K As Integer = 1 To 3
            For J As Integer = 1 To 3
                If m_lstBC(iNumAss1).asMot(J) = "" Then GoTo MotSuivant
                If m_lstBC(iNumAss2).asMot(K) <> m_lstBC(iNumAss1).asMot(J) Then GoTo MotSuivant

                sMotPivot = m_lstBC(iNumAss2).asMot(K)
                bMotTrouve = True
                bAucunMotTrouve = False
                iPosMT_PMaj = J : iPosMT_PMin = K
                J = iNbMotsBCMax : K = iNbMotsBCMax ' Fin de boucle

MotSuivant:
            Next J
        Next K

        If bRechecheApprofondie And Not bMotTrouve Then GoTo AssertionPrecedente
        If Not bRechecheApprofondie And Not bMotTrouve Then
            ReponseIAVB(sReponseRienConclure, bAffiche2Etoiles:=True)
            Exit Sub
        End If

        Dim bInversion As Boolean = False
        Dim bFctPetitTerme As Boolean = False
        Dim bFctGdTerme As Boolean = False
        Dim sPetitTerme$ = ""

        ' Résolution
        Select Case iPosMT_PMin
        Case 1
            sPetitTerme = sVBRight(m_lstBC(iNumAss2).sEntree, _
                iLenA2 - m_lstBC(iNumAss2).iPosFinMot1)
            ' Ex.: EST PHILOSOPHE
            If m_bVersionModifiee And sVBLeft(sPetitTerme, 4) = sMotEst & sEspace Then _
                bInversion = True

        Case 2
            sPetitTerme = sVBLeft(m_lstBC(iNumAss2).sEntree, m_lstBC(iNumAss2).iPosFinMot1)
            ' Ex.: "OR SOCRATE "

        Case 3
            sPetitTerme = sVBLeft(m_lstBC(iNumAss2).sEntree, m_lstBC(iNumAss2).iPosFinMot2)
            If m_bVersionModifiee And sVBRight(sPetitTerme, 1) = sParenthOuv Then _
                bFctPetitTerme = True

        End Select

        ' Ex.: OR TOUT HOMME SENSE
        If sVBLeft(sPetitTerme, 3) = sMotOr & sEspace Then sPetitTerme = _
            sVBRight(sPetitTerme, iVBLen(sPetitTerme) - 3)

        Dim sGdTerme$ = ""
        Select Case iPosMT_PMaj
        Case 1
            sGdTerme = sVBRight(m_lstBC(iNumAss1).sEntree, iLenA1 - m_lstBC(iNumAss1).iPosFinMot1)
            ' Ex.: "EST MORTEL"
            ' Ex.: "EST GREC"
            ' Ex.: "EST INCOMPRIS"
            If bFctPetitTerme And sVBLeft(sGdTerme, 4) = sMotEst & sEspace Then _
                sGdTerme = sVBTrim(sVBRight(sGdTerme, iVBLen(sGdTerme) - 4)) & sEspace & sParenthFerm

        Case 2
            Syllogisme2(iNumAss1, bFctPetitTerme, bInversion, sGdTerme)

        Case 3
            sGdTerme = sVBLeft(m_lstBC(iNumAss1).sEntree, m_lstBC(iNumAss1).iPosFinMot2)
            If m_bVersionModifiee And sVBRight(sGdTerme, 1) = sParenthOuv Then _
                bFctGdTerme = True

        End Select

        SyllogismeConclusion(iPosMT_PMin, iPosMT_PMaj, sPetitTerme, sGdTerme, _
            bInversion, bFctGdTerme, sMotPivot)

AssertionPrecedente:
    Next iNumAss1

    If bRechecheApprofondie And Not bPasse2 Then bPasse2 = True : GoTo Passe2

    ReponseIAVB("", bFinListe:=True)

    If bAucunMotTrouve Then ReponseIAVB(sReponseRienConclure, bAffiche2Etoiles:=True)

End Sub

Private Sub Syllogisme2(iNumAss1%, bFctPetitTerme As Boolean, bInversion As Boolean, _
    ByRef sGdTerme$)

    sGdTerme = sVBLeft(m_lstBC(iNumAss1).sEntree, m_lstBC(iNumAss1).iPosFinMot1)

    If Not m_bVersionModifiee Then Exit Sub

    If sVBLeft(sGdTerme, 3) = sMotOr & sEspace Then
        If bFctPetitTerme Then _
            sGdTerme = sVBTrim(sVBRight(sGdTerme, iVBLen(sGdTerme) - 3)) & sEspace & sParenthFerm : Exit Sub
        If bInversion Then
            sGdTerme = sVBTrim(sVBRight(sGdTerme, iVBLen(sGdTerme) - 3))
        Else
            sGdTerme = sMotEstComme & sEspace & _
                sVBTrim(sVBRight(sGdTerme, iVBLen(sGdTerme) - 3))
        End If
        Exit Sub
    End If

    If bFctPetitTerme Then sGdTerme = sVBTrim(sGdTerme) & sEspace & sParenthFerm : Exit Sub

    If sVBLeft(sGdTerme, 3) <> sMotEst AndAlso sVBLeft(sGdTerme, 4) <> sMotSont AndAlso _
        (sVBLeft(sGdTerme, 3) <> sMotTou OrElse Not bInversion) Then _
        sGdTerme = sMotEstComme & sEspace & sVBTrim(sGdTerme)

End Sub

Private Sub SyllogismeConclusion(iPosMT_PMin%, iPosMT_PMaj%, sPetitTerme$, sGdTerme$, _
    ByRef bInversion As Boolean, bFctGdTerme As Boolean, sMotPivot$)

    If iPosMT_PMin + iPosMT_PMaj = 2 Then
        If sVBLeft(sPetitTerme, 4) = sMotEst & sEspace Then
            sPetitTerme = sVBRight(sPetitTerme, iVBLen(sPetitTerme) - 4)
            If sVBLeft(sPetitTerme, 3) = sMotUn & sEspace Then _
                sPetitTerme = sVBRight(sPetitTerme, iVBLen(sPetitTerme) - 3)
            sPetitTerme = sMotQuelque & sEspace & sVBTrim(sPetitTerme) & sEspace
            bInversion = False
        ElseIf sVBLeft(sPetitTerme, 5) = sMotSont & sEspace Then
            sPetitTerme = sVBRight(sPetitTerme, iVBLen(sPetitTerme) - 5)
            sPetitTerme = sMotQuelques & sEspace & sVBTrim(sPetitTerme)
            bInversion = False
        End If
    End If

    sPetitTerme = sVBTrim(sPetitTerme)
    sGdTerme = sVBTrim(sGdTerme)

    If m_bVersionModifiee Then ' 08/04/2017
        If bInversion Then
            If sVBLeft(sGdTerme, 1) = sEgal Then sGdTerme = sVBMid(sGdTerme, 2)
        Else
            If sVBLeft(sPetitTerme, 1) = sEgal Then sPetitTerme = sVBMid(sPetitTerme, 2)
        End If
    End If

    Dim sConclusion$ = sMotDonc & sEspace & sPetitTerme & sEspace & sGdTerme
    If bInversion Then sConclusion = sMotDonc & sEspace & sGdTerme & sEspace & sPetitTerme
    If bFctGdTerme Then
        If bInversion Then sPetitTerme = sVBRight(sPetitTerme, iVBLen(sPetitTerme) - 4)
        sConclusion = sMotDonc & sEspace & sGdTerme & sEspace & sPetitTerme & sEspace & sParenthFerm
    End If
    Dim sConclusionParlee$ = sConclusion

    If m_bVersionModifiee Then sConclusion = sConclusion & "  <" & sMotPivot & ">"

    ReponseIAVB(sConclusion, sTexteParleSpecifique:=sConclusionParlee, bListe:=True)

    ' Ex. de syllogisme :
    ' TOUT HOMME EST MORTEL
    ' OR SOCRATE EST UN HOMME
    ' DONC SOCRATE EST MORTEL

    ' Il y a plusieurs formes de syllogisme, la conclusion peut en varier :
    ' PLATON EST GREC
    ' OR PLATON EST PHILOSOPHE
    ' DONC QUELQUE PHILOSOPHE EST GREC

    ' Dans certains cas, un syllogisme peut être trouvé avec 2 mots au lieu d'un seul :
    ' TOUT LOGICIEN EST INCOMPRIS
    ' OR TOUT HOMME SENSE EST LOGICIEN
    ' DONC TOUT HOMME SENSE EST INCOMPRIS

End Sub

#End Region

#Region "Initialisations"

Private Sub InitCmd()

    sCmdEffBase = "/EFF"  ' Effacer toute la base
    sCmdEff = "/D"        ' Effacer la denière assertion, ou /Dn : l'assertion n° n
    sCmdLister = "/L"     ' Lister les assertions de la base
    sCmdCopier = "/C"     ' Copier la discussion dans le presse papier
    sCmdSilence = "/S"    ' Mettre en sourdine
    sCmdParler = "/PARLE" ' Ré-activer la voix

    sMotQU = "QU"
    sMotDonc = "DONC"
    sMotEstComme = "EST COMME"
    sMotEst = "EST"
    sMotSont = "SONT"
    sMotTou = "TOU" '(TOUT, TOUS, TOUTES)
    sMotUn = "UN"
    sMotOr = "OR"
    sMotQuelque = "QUELQUE"
    sMotQuelques = "QUELQUES"

    If m_bTraiterEnMinuscules Then

        sCmdEffBase = sCmdEffBase.ToLowerInvariant()
        sCmdEff = sCmdEff.ToLowerInvariant()
        sCmdLister = sCmdLister.ToLowerInvariant()
        sCmdCopier = sCmdCopier.ToLowerInvariant()
        sCmdSilence = sCmdSilence.ToLowerInvariant()
        sCmdParler = sCmdParler.ToLowerInvariant()

        sMotQU = sMotQU.ToLowerInvariant()
        sMotDonc = sMotDonc.ToLowerInvariant()
        sMotEstComme = sMotEstComme.ToLowerInvariant()
        sMotEst = sMotEst.ToLowerInvariant()
        sMotSont = sMotSont.ToLowerInvariant()
        sMotTou = sMotTou.ToLowerInvariant()
        sMotUn = sMotUn.ToLowerInvariant()
        sMotOr = sMotOr.ToLowerInvariant()
        sMotQuelque = sMotQuelque.ToLowerInvariant()
        sMotQuelques = sMotQuelques.ToLowerInvariant()

    End If

End Sub

Private Sub InitMotsIgnores()

    ' Mots ignorés (car non-signifiants)

    ' "ETE" "ETAIT" "CA"
    m_lstMotsIgnores = New List(Of String) From { _
        "EST", "LE", "LA", "DE", "UN", "UNE", "OR", "L", "DU", "D",
        "LES", "DES", "ET", "QU", "QUE", "QUI", "SONT", "IL", "ELLE",
        "A", "T", "ÉTÉ", "EN", "OU", "COMMENT", "AU", "N", "NE", "S", "SE", "ÉTAIT",
        "QUOI", "C", "CE", "QUEL", "QUELLE", "QUELS", "QUELLES",
        "PAR", "LEQUEL", "LAQUELLE", "ÇA", "SIGNIFIE",
        "TOUT", "TOUS", "TOUTE", "TOUTES", "(", ")", "=", "Y"}

    ' 29/04/2017
    If m_bTraiterSansAccents Then
        For i As Integer = 0 To m_lstMotsIgnores.Count - 1
            m_lstMotsIgnores(i) = sEnleverAccents(m_lstMotsIgnores(i))
        Next
    End If

    If Not m_bTraiterEnMinuscules Then Exit Sub
    ' Convertir tout en minuscules
    For i As Integer = 0 To m_lstMotsIgnores.Count - 1
        m_lstMotsIgnores(i) = m_lstMotsIgnores(i).ToLowerInvariant()
    Next

End Sub

Private Shared Function bQuestionResponseOui(sQuestion$) As Boolean

    Dim dlgResult As Windows.Forms.DialogResult = _
        Windows.Forms.MessageBox.Show(sQuestion, sTitreMsg, _
        Windows.Forms.MessageBoxButtons.YesNo, _
        Windows.Forms.MessageBoxIcon.Question)
    If dlgResult = Windows.Forms.DialogResult.No Then Return False
    Return True

End Function

#End Region

End Class

End Namespace