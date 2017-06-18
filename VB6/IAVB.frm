VERSION 5.00
Begin VB.Form FrmIAVB 
   Caption         =   "Intelligence Artificielle en Visual Basic"
   ClientHeight    =   6015
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7590
   Icon            =   "IAVB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSuivant 
      Caption         =   "->"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Exemple suivant"
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton CmdGo 
      Caption         =   "!"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      ToolTipText     =   "Exécution (ou entrée)"
      Top             =   3600
      Width           =   495
   End
   Begin VB.ListBox ListAssert 
      Height          =   1425
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Exemples à tester (Click ou DblClick)"
      Top             =   4440
      Width           =   7335
   End
   Begin VB.ListBox ListIA 
      BackColor       =   &H00FFFFFF&
      Height          =   2985
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "IAVB"
      Top             =   120
      Width           =   7335
   End
   Begin VB.TextBox TextInput 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Zone de saisie"
      Top             =   3480
      Width           =   5895
   End
End
Attribute VB_Name = "FrmIAVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Fichier frmIAVB.vb : Intelligence Artificielle en Visual Basic (IAVB)
' ------------------

' http://www.vbfrance.com/code.aspx?ID=1860
' Documentation : IAVB.html
' http://patrice.dargenton.free.fr/ia/iavb/index.html
' http://patrice.dargenton.free.fr/ia/iavb/IAVB.vbproj.html
' Version 2.11 du 18/06/2017

' IAVB est la transcription en Visual Basic d'un logiciel paru dans la revue
'  MICRO-SYSTEMES en Décembre 1984, pages 195-202 :
'  "Mini-système expert pour Apple II" par Philippe LARVET :
'  Gestion d'une base de connaissances.

' Par Patrice Dargenton : mailto:patrice.dargenton@free.fr
' http://patrice.dargenton.free.fr/index.html
' http://patrice.dargenton.free.fr/CodesSources/index.html
' http://www.vbfrance.com/listeauteur2.aspx?ID=1124

' Conventions de nommage des variables :
' ------------------------------------
' b pour Boolean (booléen vrai ou faux)
' i pour Integer : % (en VB .Net, l'entier a la capacité du VB6.Long)
' l pour Long : &
' r pour nombre Réel (Single!, Double# ou Decimal : D)
' s pour String : $
' c pour Char ou Byte
' d pour Date
' a pour Array (tableau) : ()
' o pour Object : objet instancié localement
' refX pour reference à un objet X préexistant qui n'est pas sensé être fermé
' m_ pour variable Membre de la classe ou de la feuille (Form)
'  (mais pas pour les constantes)
' frm pour Form
' cls pour Classe
' mod pour Module
' ...
' ------------------------------------
    
Option Explicit

' Booléen pour forcer le fonctionnement dans la version originale (1984)
'  ou bien modifiée (2001)
Const bVersionModifiee As Boolean = True
'Const bVersionModifiee As Boolean = False

' Nbre de mots signifiants de l'entrée max.
' Peut être > iNbMotsBCMax notamment pour les compositions de fonctions
Const iNbMotsEntreeMax% = 10

Const sGM = """"
Const iCodeToucheEntree% = 13
'Const sSilence$ = "-"

' Pour se positionner sur la dernière ligne de la liste des assertions
Dim bPositionnement As Boolean
Const iNumLigneDebutExemples% = 8 ' Pour la démo sur les exemples

' Mots ignorés (car non-signifiants) de la table asMotsIgnores
Const iNbMotsIgnores% = 51

' A faire en dynamique :
' Nbre d'assertions max.
Const iNbAssertionsMax% = 100

' iNbMotsBCMax : Nbre de mots signifiants susceptibles
'  d'être conservé en mémoire pour chaque assertion (au moins 3).
'  seul les 2 premiers mots peuvent former une clé pour une relation H ou une indirection
Const iNbMotsBCMax% = 4

Private Type TAssertion
    iNbMots As Integer
    asMot(iNbMotsBCMax) As String
    iPosFinMot1 As Integer
    iPosFinMot2 As Integer
    sEntree As String
End Type

Private Type TParamEntree
    bMotQuestion      As Boolean
    bInterrogation    As Boolean
    bFinTraitement    As Boolean
    sEntree           As String
    ' sEntreeCompilee : Concaténation de tous les asMotsEntree$(i)
    sEntreeCompilee   As String
    ' iNbMotsEntree : Nbre de mots signifiants extraits de sEntree
    iNbMotsEntree     As Integer
    ' asMotsEntree : Mots signifiants extraits de l'assertion
    asMotsEntree(iNbMotsEntreeMax) As String
    iPosFinMot1       As Integer
    iPosFinMot2       As Integer
    
    ' n° de l'assertion contenant le mot signifiant n°K de l'entrée,
    '  tableau des K mots de l'entrée
    '  D'abord la première assertion : ControleExistenceMots(),
    '  puis les assertions suivantes : bAssertCMot() pour les indirections.
    aiNumAssertCMot(iNbMotsEntreeMax) As Integer
    
    ' Conservation de l'indice minimum pour
    '  l'espace de recherche max. de l'assertion
    '  (contenant un des mots de l'assertion en cours)
    iNumAssertMinRech As Integer

End Type

Private Type TGlobal
    ' Table des mots ignorés en général (non-signifiants)
    asMotsIgnores(iNbMotsIgnores) As String
    
    bQuestionPreced_Donc As Boolean
    
    iNbAssertions As Integer ' Nbre d'assertions stockées dans la base aBC
    aBC(iNbAssertionsMax) As TAssertion ' Base de Connaissances (KB)
    
    iNumAssertionEnCours As Integer
    
    bVoixActive As Boolean
    
    bAgent As Boolean
    
End Type
    
Private glb As TGlobal

' Moins de 1% du code est distinct entre les versions VBA et VB6
'  en utilisant la compilation conditionnelle, on ne maintient qu'une seule version !
#If bVersionVBA Then ' Version VBA (Visual Basic pour Application) Excel et Word

Private Sub UserForm_Initialize()
    Initialiser
End Sub
Private Sub UserForm_Activate()
    Activer
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'If Not bQuitter() Then Cancel = True
End Sub

#Else ' Version VB6

Private Sub Form_Load()
    Initialiser
End Sub
Private Sub Form_Activate()
    Activer
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'If Not bQuitter() Then Cancel = True
End Sub

#End If

Private Sub Activer()
    
End Sub

Private Sub ReponseIAVB(sTexte$, _
    Optional ByVal bAffiche2Etoiles As Boolean = False, _
    Optional ByVal bTexteAfficherAvant As Boolean = False, _
    Optional ByVal bSilence As Boolean = False, _
    Optional ByVal sTexteParleSpecifique$ = "")
    
    Dim sTexteEcrit$
    sTexteEcrit = sTexte
    If bAffiche2Etoiles Then sTexteEcrit = "** " & sTexte
    
    ' Affichage du texte avant de parler
    If bTexteAfficherAvant Then ListIA.AddItem sTexteEcrit
    
    If glb.bVoixActive Then
        Dim sTexteParle$
        sTexteParle = sTexte
        If Len(sTexteParleSpecifique) > 0 Then sTexteParle = sTexteParleSpecifique
        
        ' Toujours se positionner sur la dernière ligne
        PositionnerListIA
        
        If bSilence Then GoTo Suite
        
        If glb.bAgent Then
            ' Si les minuscules améliore la prononciation alors dans ce cas :
            sTexteParle = sConvMinuscules(sTexteParle)
            bDire sTexteParle
        End If
    Else
        PositionnerListIA
    End If
    
Suite:
    If Not bTexteAfficherAvant Then ListIA.AddItem sTexteEcrit

End Sub
    
Public Function bDire(sParole$) As Boolean
    
    If Not glb.bAgent Then Exit Function
    If Not glb.bVoixActive Then Exit Function
    Dim bOk As Boolean
    
#If bVersionVBA Then
    ' Pas de synthèse vocale
#Else ' Version VB6
    bOk = bDireMSAgent(sParole)
#End If
    
    bDire = bOk

End Function
    
Private Function sConvMinuscules$(sTexteParle$)

    ' Première lettre en maj. et min. ensuite
    sConvMinuscules = UCase$(Left$(sTexteParle, 1)) & LCase$(Mid$(sTexteParle, 2))

End Function
    
Private Sub IAVBMain()
    
    Dim prm As TParamEntree
    AnalyseEntree prm
    If prm.bFinTraitement Then Exit Sub
    
    If prm.bInterrogation And prm.asMotsEntree(1) = "DONC" Then
        Dim bRechecheApprofondie As Boolean
        bRechecheApprofondie = False
        If bVersionModifiee And glb.bQuestionPreced_Donc Then
            bRechecheApprofondie = True
            If glb.iNumAssertionEnCours > 0 Then _
                ReponseIAVB glb.aBC(glb.iNumAssertionEnCours).sEntree & " :"
            ReponseIAVB "Recherche approfondie...", bAffiche2Etoiles:=True
        End If
        Syllogisme prm, bRechecheApprofondie
        glb.bQuestionPreced_Donc = True
        Exit Sub
    End If
    glb.bQuestionPreced_Donc = False
    
    If Not prm.bInterrogation Then
        If prm.iNbMotsEntree <= 1 Then
            ReponseIAVB "Votre phrase est trop courte !", bAffiche2Etoiles:=True
            Exit Sub
        End If
    
        ' Assertion : Controle existence de l'assertion dans la base
        AjoutBase prm: Exit Sub
    End If

    ' Interrogation : controle existence des mots
    ControleExistenceMots prm
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
        ReponseIAVB "Je l'ignore.", bAffiche2Etoiles:=True
        Exit Sub
    End If
    ReponseIAVB "Non.", bAffiche2Etoiles:=True

End Sub

Private Sub AnalyseEntree(prm As TParamEntree)

    prm.bInterrogation = False
    prm.bInterrogation = False
    prm.bFinTraitement = False
    prm.bMotQuestion = False
    prm.iNbMotsEntree = 0
    prm.sEntreeCompilee = ""
    
    Dim I%, J%, K%
    Dim iPos%, iMemPos% ' Boucles sur les lettres
    Dim sLettreEntree$, iLenEntree%
    Dim sMotEntree$ ' Chaque mot extrait de sEntree
    
    ' Extraction des mots
    For K = 1 To iNbMotsBCMax
        prm.asMotsEntree(K) = "": prm.aiNumAssertCMot(K) = 0
    Next K
    prm.sEntree = TextInput.Text
    iLenEntree = Len(prm.sEntree)
    
    ' Examen de l'entrée
    If Left$(prm.sEntree, 1) = "" Then
        ReponseIAVB prm.sEntree, bSilence:=True
        prm.bFinTraitement = True: Exit Sub
    End If
    ' Gestion des commentaires
    If Left$(prm.sEntree, 1) = "'" Then
        ReponseIAVB prm.sEntree, bSilence:=True
        prm.bFinTraitement = True: Exit Sub
    End If
    If Left$(prm.sEntree, 1) = "/" Then
        ReponseIAVB prm.sEntree, bSilence:=True
        GoTo TraitementCmd
    End If
    
    ' Afficher immédiatement le texte, sans attendre qu'il soit prononcé
    ReponseIAVB prm.sEntree, bTexteAfficherAvant:=True
    
    If Right$(prm.sEntree, 1) = "?" Then
        prm.sEntree = Left$(prm.sEntree, iLenEntree - 1)
        iLenEntree = Len(prm.sEntree)
        prm.bInterrogation = True
    End If
    
    iPos = 1

    Do
        iMemPos = iPos
    
        Do
            iPos = iPos + 1
            sLettreEntree = Mid$(prm.sEntree, iPos, 1)
        Loop While iPos <= iLenEntree And _
           sLettreEntree <> " " And sLettreEntree <> "'" And _
           sLettreEntree <> "-" And sLettreEntree <> "(" And sLettreEntree <> ")"
        
        sMotEntree = Mid$(prm.sEntree, iMemPos, iPos - iMemPos)
        If iMemPos = 1 And Left$(sMotEntree, 2) = "QU" Then _
            prm.bMotQuestion = True: GoTo RechercheMotSuivant
        
        Dim bMotIgnore As Boolean
        bMotIgnore = False
        For K = 1 To iNbMotsIgnores
            If glb.asMotsIgnores(K) = sMotEntree Then bMotIgnore = True: Exit For
        Next K
        If bMotIgnore Then GoTo RechercheMotSuivant
        
        ' Ajout d'un mot signifiant
        prm.iNbMotsEntree = prm.iNbMotsEntree + 1
        ' Modification : test aussi avec iNbMotsEntreeMax
        If (Not prm.bInterrogation And prm.iNbMotsEntree > iNbMotsBCMax) Or _
           prm.iNbMotsEntree > iNbMotsEntreeMax Then
            ReponseIAVB "Votre phrase est trop longue !", bAffiche2Etoiles:=True
            prm.bFinTraitement = True: Exit Sub
        End If
            
        prm.sEntreeCompilee = prm.sEntreeCompilee + sMotEntree
        prm.asMotsEntree(prm.iNbMotsEntree) = sMotEntree
        If prm.iNbMotsEntree = 1 Then prm.iPosFinMot1 = iPos
        If prm.iNbMotsEntree = 2 Then prm.iPosFinMot2 = iPos
    
RechercheMotSuivant:
        If iPos > iLenEntree Then Exit Sub
        
        Do
            iPos = iPos + 1
        Loop While Mid$(prm.sEntree, iPos, 1) = " " And iPos <= iLenEntree
        
    Loop While iPos <= iLenEntree
    
    If prm.asMotsEntree(1) = "" Then
        ReponseIAVB "Votre phrase est incomplète !", bAffiche2Etoiles:=True
            
        prm.bFinTraitement = True: Exit Sub
    End If

    Exit Sub
    
TraitementCmd:
    If Left$(prm.sEntree, 2) = "/L" Then GoTo ListageBaseC
    If Left$(prm.sEntree, 4) = "/EFF" Then GoTo EffacementBase
    If Left$(prm.sEntree, 2) = "/D" Then GoTo Supression
    If Left$(prm.sEntree, 2) = "/C" Then GoTo CopiePressePapier
    
    If Left$(prm.sEntree, 2) = "/S" Then
        ReponseIAVB "Okay, je la ferme.", bAffiche2Etoiles:=True
        glb.bVoixActive = False
        prm.bFinTraitement = True: Exit Sub
    End If
    
    If Left$(prm.sEntree, 6) = "/PARLE" Then
        glb.bVoixActive = glb.bAgent
        ReponseIAVB "Je vous écoute.", bAffiche2Etoiles:=True
        prm.bFinTraitement = True: Exit Sub
    End If
    
    ReponseIAVB "Commande inconnue !", bAffiche2Etoiles:=True
    prm.bFinTraitement = True: Exit Sub

ListageBaseC:
    If glb.iNbAssertions = 0 Then ReponseIAVB "Base vide.", bAffiche2Etoiles:=True
    For I = 1 To glb.iNbAssertions
        ListIA.AddItem (I & " " & glb.aBC(I).sEntree)
    Next I
    prm.bFinTraitement = True: Exit Sub
    
CopiePressePapier:
    ' Copie du résultat d'exécution dans le presse papier
    Dim sListIA$
    For I = 0 To ListIA.ListCount - 1
        sListIA = sListIA & ListIA.List(I) & vbCrLf
    Next I
    CopierPressePapier sListIA
    ReponseIAVB "Le résultat d'exécution a été copié dans le presse papier.", _
        bAffiche2Etoiles:=True
        
    prm.bFinTraitement = True: Exit Sub

Supression:
    I = glb.iNbAssertions
    If I = 0 Then
        ReponseIAVB "Base vide.", bAffiche2Etoiles:=True
        prm.bFinTraitement = True: Exit Sub
    End If
    
    'If iLenEntree > 2 Then I = Val(Right$(prm.sEntree, iLenEntree - 2))
    If Mid$(prm.sEntree, 3, 1) <> " " Then
        I = Val(Right$(prm.sEntree, iLenEntree - 2))
    End If
    For J = 1 To iNbMotsBCMax
        glb.aBC(I).asMot(J) = ""
    Next J
    glb.aBC(I).sEntree = ""

    ReponseIAVB "Assertion " & I & " Supprimée.", bAffiche2Etoiles:=True
        
    prm.bFinTraitement = True: Exit Sub

EffacementBase:
    If glb.iNbAssertions = 0 Then
        ReponseIAVB "Base vide.", bAffiche2Etoiles:=True
        prm.bFinTraitement = True: Exit Sub
    End If
    ReponseIAVB "Voulez-vous effacer toute la base ?", bAffiche2Etoiles:=True
    If vbNo = MsgBox("** Voulez-vous effacer toute la base ?", _
        vbQuestion Or vbYesNo, sTitreMsg) Then
        prm.bFinTraitement = True: Exit Sub
    End If
    For I = 1 To glb.iNbAssertions
        For J = 1 To iNbMotsBCMax
            glb.aBC(I).asMot(J) = ""
        Next J
    Next I
    ' Les accents font planter le contrôle TextToSpeech !!!
    ReponseIAVB "Base entièrement effacée !", bAffiche2Etoiles:=True
                
    glb.iNbAssertions = 0
    prm.bFinTraitement = True
    
End Sub

Private Sub AjoutBase(prm As TParamEntree)
    
    ' Assertion : Controle existence de l'assertion dans la base
    Dim I%, J%, sGdTerme$
    For I = 1 To glb.iNbAssertions
        sGdTerme = ""
        For J = 1 To iNbMotsBCMax
           sGdTerme = sGdTerme + glb.aBC(I).asMot(J)
        Next J
        If sGdTerme = prm.sEntreeCompilee Then
            If bVersionModifiee Then glb.iNumAssertionEnCours = I
            ReponseIAVB "Assertion déjà connue !", bAffiche2Etoiles:=True
            Exit Sub
        End If
    Next I
    
    ' Enrichissement base
    Dim bPlaceLibre As Boolean
    bPlaceLibre = False
    For I = 1 To iNbAssertionsMax
        If glb.aBC(I).asMot(1) = "" Then
            ' Il reste de la place dans la base
            bPlaceLibre = True: Exit For
        End If
    Next I

    If Not bPlaceLibre Then
        ReponseIAVB "Stop ! La base est pleine !", bAffiche2Etoiles:=True
        Exit Sub
    End If
    
    For J = 1 To iNbMotsBCMax
        glb.aBC(I).asMot(J) = prm.asMotsEntree(J)
    Next J
    glb.aBC(I).sEntree = prm.sEntree
    glb.aBC(I).iNbMots = prm.iNbMotsEntree
    glb.aBC(I).iPosFinMot1 = prm.iPosFinMot1
    glb.aBC(I).iPosFinMot2 = prm.iPosFinMot2
    ReponseIAVB "Compris.", bAffiche2Etoiles:=True
    If I > glb.iNbAssertions Then
        glb.iNbAssertions = I
        glb.iNumAssertionEnCours = glb.iNbAssertions
    End If

End Sub

Private Sub ControleExistenceMots(prm As TParamEntree)
    
    ' Interrogation : controle existence des mots
    
    Dim I%, J%, K%, bMotPresentAssertion As Boolean, bMotTrouve As Boolean
    Dim iLenEntree%, iMemNumAssertionEnCours%
    glb.iNumAssertionEnCours = glb.iNbAssertions
    iLenEntree = Len(prm.sEntree)
    prm.iNumAssertMinRech = glb.iNbAssertions ' Util pour la version modifiée
    
    For K = 1 To prm.iNbMotsEntree
        bMotPresentAssertion = False
        For I = 1 To glb.iNbAssertions
            bMotTrouve = False
            For J = 1 To glb.aBC(I).iNbMots
                If glb.aBC(I).asMot(J) = prm.asMotsEntree(K) Then bMotTrouve = True: Exit For
            Next J
            If Not bMotTrouve Then GoTo AssertionSuivante
            
            bMotPresentAssertion = True
            If prm.aiNumAssertCMot(K) = 0 Then prm.aiNumAssertCMot(K) = I
            
            If prm.iNbMotsEntree > 1 Then iMemNumAssertionEnCours = I: Exit For
            
            If J = 2 And iLenEntree > Len(prm.asMotsEntree(K)) + 1 Then
                ReponseIAVB Left$(glb.aBC(I).sEntree, glb.aBC(I).iPosFinMot1)
                ' Question spécifique, ex.:
                ' ET LAQUELLE EST FOFOLLE ?
                ' MINNA = Left$(aBC(I).sEntree, aBC(I).iPosFinMot1)
                ' aBC(I).sEntree = MINNA EST FOFOLLE
                GoTo AssertionSuivante
            End If
            ReponseIAVB glb.aBC(I).sEntree
            ' Entrée contenant le mot K : Question générique, ex.:
            ' RESPONSABLE ?
            ' M.BERTRAND EST RESPONSABLE DE L'ANNEXE
            ' M.JACQUES EST RESPONSABLE DE LA SAISIE
    
AssertionSuivante:
        Next I
        
        If Not bMotPresentAssertion Then
            ReponseIAVB "Je ne connais pas " & prm.asMotsEntree(K), bAffiche2Etoiles:=True
            prm.bFinTraitement = True: Exit Sub
        End If
        If prm.iNbMotsEntree = 1 Then prm.bFinTraitement = True: Exit Sub
        
        ' Conserver l'indice minimum pour l'espace de recherche max de l'assertion
        If bVersionModifiee Then
            If iMemNumAssertionEnCours < prm.iNumAssertMinRech Then _
                prm.iNumAssertMinRech = iMemNumAssertionEnCours
        Else
            If iMemNumAssertionEnCours < glb.iNumAssertionEnCours Then _
                glb.iNumAssertionEnCours = iMemNumAssertionEnCours
        End If
            
    Next K
    
End Sub

Private Function bRelationHorizontale(prm As TParamEntree) As Boolean

    Dim I%, J%, K%, iNbMotsTrouves%, iDeb%
    Dim aiMotsTrouves%(2)
    bRelationHorizontale = False
    iDeb = glb.iNumAssertionEnCours
    If bVersionModifiee Then iDeb = prm.iNumAssertMinRech
    
    For I = iDeb To glb.iNbAssertions
        iNbMotsTrouves = 0
        Dim iMemJ%
        iMemJ = 1

        For K = 1 To prm.iNbMotsEntree
            Dim bMotTrouve As Boolean
            bMotTrouve = False
            For J = iMemJ To glb.aBC(I).iNbMots
                If glb.aBC(I).asMot(J) = prm.asMotsEntree(K) Then bMotTrouve = True: Exit For
            Next J
            If bMotTrouve Then
                iNbMotsTrouves = iNbMotsTrouves + 1
                If iNbMotsTrouves <= 2 Then aiMotsTrouves(iNbMotsTrouves) = J
            End If
            iMemJ = J ' Force l'ordre de recherche à Sujet/Verbe/Complement seulement
            ' Cela permet de répondre Oui à la question :
            ' SABINE AIME-T-ELLE JACQUES ?
            ' et cela empêche de trouver une relation H à la question :
            ' JACQUES AIME-T-IL SABINE ?
        Next K

        If iNbMotsTrouves = 0 Then GoTo AssertionSuivante
        
        If iNbMotsTrouves = prm.iNbMotsEntree And _
           iNbMotsTrouves = glb.aBC(I).iNbMots Then
            ReponseIAVB "Oui.", bAffiche2Etoiles:=True
            bRelationHorizontale = True:  Exit Function
        End If
        
        If iNbMotsTrouves <> prm.iNbMotsEntree Then GoTo AssertionSuivante
        bRelationHorizontale = True
        
        If aiMotsTrouves(1) = 1 And aiMotsTrouves(2) = 2 Then
            Dim iLenE%
            iLenE = Len(glb.aBC(I).sEntree)
            ReponseIAVB Right$(glb.aBC(I).sEntree, iLenE - glb.aBC(I).iPosFinMot2)
            ' Ex.:
            ' QUI MARIE REGARDE-T-ELLE ?
            ' Réponse : HUGUES
            ' aBC(I).sEntree = MARIE REGARDE HUGUES
            GoTo AssertionSuivante
        End If
        
        If aiMotsTrouves(1) = 2 Then
            ReponseIAVB Left$(glb.aBC(I).sEntree, glb.aBC(I).iPosFinMot1)
            ' Ex.:
            ' CHEF DE SERVICE ?
            ' M.RENE
            ' M.DUBOIS
            ' aBC(I).sEntree = M.RENE EST LE CHEF DU SERVICE COMPTABILITE
            ' aBC(I).sEntree = M.DUBOIS EST CHEF DU SERVICE PHOTO
            GoTo AssertionSuivante
        End If
        
        If aiMotsTrouves(1) = 3 Then
            ReponseIAVB Left$(glb.aBC(I).sEntree, glb.aBC(I).iPosFinMot2)
            ' Ex.:
            ' SERVICE PHOTO ?
            ' M.DUBOIS EST CHEF
            GoTo AssertionSuivante
        End If
        
        ReponseIAVB glb.aBC(I).sEntree
        ' Pas d'exemple !

AssertionSuivante:
    Next I
    
End Function

Private Function bComposFonction(prm As TParamEntree) As Boolean
    
    ' Composition de fonction (relation verticale)
    
    bComposFonction = False
    Dim I%, K%, asMotReduit$(iNbMotsEntreeMax)
    For K = 1 To prm.iNbMotsEntree
        asMotReduit(K) = prm.asMotsEntree(K)
    Next K
    
    For K = prm.iNbMotsEntree To 2 Step -1
        Dim sGdTerme$, bCompFct As Boolean, iDeb%
        sGdTerme = asMotReduit(K - 1) + asMotReduit(K)
        bCompFct = False
        
        ' Recherche des gd termes sur les mots n° 2 et 3 seulement de la BC :
        '  du type Y = F(X) : Y:1, F:2 et X:3, par ex.: BLANC = COULEUR(ARTABAN)
        iDeb = glb.iNumAssertionEnCours
        If bVersionModifiee Then iDeb = prm.iNumAssertMinRech
        For I = iDeb To glb.iNbAssertions
            If glb.aBC(I).asMot(2) + glb.aBC(I).asMot(3) = sGdTerme Then _
                bCompFct = True: Exit For
        Next I
        If Not bCompFct Then Exit Function
        
        asMotReduit(K - 1) = glb.aBC(I).asMot(1)
    Next K
    
    ' Fin de composition de fct : réponse trouvée
    ReponseIAVB asMotReduit(1)
    bComposFonction = True
    ' Ex.:
    ' ARTABAN = CHEVAL(HENRI_IV)
    ' BLANC = COULEUR(ARTABAN)
    ' HENRI_IV = ROI(NAVARRE)
    ' COULEUR(CHEVAL(ROI(NAVARRE))) ?
    ' Réponse : BLANC
    ' sGdTerme = COULEURARTABAN

End Function

Private Function bIndirection(prm As TParamEntree) As Boolean

    Dim iMemNumAssertCMot2%, I1%, I2%, J1%, J2%, sGdTerme$
    iMemNumAssertCMot2 = prm.aiNumAssertCMot(2)
    bIndirection = False

    Do
        prm.aiNumAssertCMot(2) = iMemNumAssertCMot2
        Do
            I1 = prm.aiNumAssertCMot(1)
            I2 = prm.aiNumAssertCMot(2)
            For J1 = 1 To glb.aBC(I1).iNbMots
                sGdTerme = glb.aBC(I1).asMot(J1)
                For J2 = 1 To glb.aBC(I2).iNbMots
                    If sGdTerme <> glb.aBC(I2).asMot(J2) Then GoTo MotSuivant
                    'Debug.Print glb.aBC(I1).sEntree
                    'Debug.Print glb.aBC(I2).sEntree
                    If I1 <> I2 And _
                        sGdTerme <> prm.asMotsEntree(1) And _
                        sGdTerme <> prm.asMotsEntree(2) And _
                       (prm.iNbMotsEntree <= 3 Or Not bVersionModifiee) Then
                       ' Modification : il ne peut y avoir indirection que sur une
                       '  question simple (au plus 3 termes signifiants)
                        bIndirection = True
                        ReponseIAVB glb.aBC(I1).asMot(J1)
                        J1 = iNbMotsBCMax: J2 = iNbMotsBCMax ' Exit For 2x
                        
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
                        ' glb.aBC(I1).asMot(J1) = ARTABAN
                        ' prm.asMotsEntree(1) = COULEUR
                        ' prm.asMotsEntree(2) = CHEVAL

                    End If
MotSuivant:
                Next J2
            Next J1
        Loop While bAssertCMot(prm, K:=2)
    Loop While bAssertCMot(prm, K:=1)

End Function

Private Function bAssertCMot(prm As TParamEntree, K%) As Boolean
    
    ' Module de recherche d'un prochain aiNumAssertCMot(K)
    Dim I%, J%, iDeb%
    iDeb = prm.aiNumAssertCMot(K)
    For I = iDeb + 1 To glb.iNbAssertions
        For J = 1 To glb.aBC(iDeb).iNbMots
            If glb.aBC(I).asMot(J) = prm.asMotsEntree(K) Then
                bAssertCMot = True
                prm.aiNumAssertCMot(K) = I
                ' Ex.:ANIMAL MANGEUR ?
                ' Réponse : CHAT, TIGRE
                ' asMotsEntree(K) = MANGEUR dans les 2 cas
                Exit Function
            End If
        Next J
    Next I

End Function

Private Sub Syllogisme(prm As TParamEntree, bRechecheApprofondie As Boolean)
    
    ' Recherche du moyen terme
    ' sPetitTerme$ : Petit terme extrait de l'assertion pour syllogisme
    ' sGdTerme$ : Concaténation de 2 ou 3 mots signifiants et extrait de l'assertion
    ' iPosMT_PMaj : Place du moyen terme dans la prémisse majeure
    ' iPosMT_PMin : Place du moyen terme dans la prémisse mineure
    Dim iNumAss1%, iNumAss2%, iLenA1%, iLenA2%, iPosMT_PMaj%, iPosMT_PMin%
    Dim sPetitTerme$, sGdTerme$
    Dim bMotTrouve As Boolean, bAucunMotTrouve As Boolean, iAssertionPreced%
    Dim bPasse2 As Boolean, iPas%, iDebut%, iFin%
    Dim sMotPivot$, J%, K%
    
    iNumAss2 = glb.iNbAssertions
    If bVersionModifiee Then _
        iNumAss2 = glb.iNumAssertionEnCours ' Maj si assertion déjà connue
    
    bAucunMotTrouve = True
    iAssertionPreced = iNumAss2 - 1
    ' Si bVersionModifiee seulement :
    If bRechecheApprofondie Then iAssertionPreced = 1 ' Recherche approfondie
    iPas = -1: iDebut = iNumAss2 - 1: iFin = iAssertionPreced

Passe2:
    If bPasse2 Then
        iPas = 1
        iDebut = iNumAss2 + 1
        iFin = glb.iNbAssertions
    End If

For iNumAss1 = iDebut To iFin Step iPas
    bMotTrouve = False
    iLenA1 = Len(glb.aBC(iNumAss1).sEntree)
    iLenA2 = Len(glb.aBC(iNumAss2).sEntree)
    For K = 1 To 3
        For J = 1 To 3
            If glb.aBC(iNumAss1).asMot(J) = "" Then GoTo MotSuivant
            If glb.aBC(iNumAss2).asMot(K) <> glb.aBC(iNumAss1).asMot(J) Then GoTo MotSuivant
            
            sMotPivot = glb.aBC(iNumAss2).asMot(K)
            bMotTrouve = True
            bAucunMotTrouve = False
            iPosMT_PMaj = J: iPosMT_PMin = K
            J = iNbMotsBCMax: K = iNbMotsBCMax ' Fin de boucle
            
MotSuivant:
        Next J
    Next K
    
    If bRechecheApprofondie And Not bMotTrouve Then GoTo AssertionPrecedente
    If Not bRechecheApprofondie And Not bMotTrouve Then
        ReponseIAVB "Je ne peux rien conclure !", bAffiche2Etoiles:=True
        Exit Sub
    End If
    
    Dim bInversion As Boolean, bFctPetitTerme As Boolean, bFctGdTerme As Boolean
    bInversion = False
    bFctPetitTerme = False
    bFctGdTerme = False

    ' Résolution
    Select Case iPosMT_PMin
    Case 1
        sPetitTerme = Right$(glb.aBC(iNumAss2).sEntree, iLenA2 - glb.aBC(iNumAss2).iPosFinMot1)
        ' Ex.: EST PHILOSOPHE
        If bVersionModifiee And Left$(sPetitTerme, 4) = "EST " Then bInversion = True
    
    Case 2
        sPetitTerme = Left$(glb.aBC(iNumAss2).sEntree, glb.aBC(iNumAss2).iPosFinMot1)
        ' Ex.: "OR SOCRATE "
    
    Case 3
        sPetitTerme = Left$(glb.aBC(iNumAss2).sEntree, glb.aBC(iNumAss2).iPosFinMot2)
        If bVersionModifiee And Right$(sPetitTerme, 1) = "(" Then
            bFctPetitTerme = True
        End If
    End Select
    
    ' Ex.: OR TOUT HOMME SENSE
    If Left$(sPetitTerme, 3) = "OR " Then sPetitTerme = Right$(sPetitTerme, Len(sPetitTerme) - 3)

    On iPosMT_PMaj GoTo Pos1, Pos2, Pos3
Pos1:
    sGdTerme = Right$(glb.aBC(iNumAss1).sEntree, iLenA1 - glb.aBC(iNumAss1).iPosFinMot1)
    ' Ex.: "EST MORTEL"
    ' Ex.: "EST GREC"
    ' Ex.: "EST INCOMPRIS"
    
    If bFctPetitTerme And Left$(sGdTerme, 4) = "EST " Then
        sGdTerme = Trim$(Right$(sGdTerme, Len(sGdTerme) - 4)) & " )"
    End If
    GoTo SuiteConclusion
    
Pos2:
    sGdTerme = Left$(glb.aBC(iNumAss1).sEntree, glb.aBC(iNumAss1).iPosFinMot1)
    
    If bVersionModifiee Then
        
        If Left$(sGdTerme, 3) = "OR " Then
            If bFctPetitTerme Then
                sGdTerme = Trim$(Right$(sGdTerme, Len(sGdTerme) - 3)) & " )"
                GoTo SuiteConclusion
            End If
            If bInversion Then
                sGdTerme = Trim$(Right$(sGdTerme, Len(sGdTerme) - 3))
            Else
                sGdTerme = "EST COMME " & Trim$(Right$(sGdTerme, Len(sGdTerme) - 3))
            End If
            GoTo SuiteConclusion
        End If
        
        If bFctPetitTerme Then
            sGdTerme = Trim$(sGdTerme) & " )"
            GoTo SuiteConclusion
        End If
        
        If Left$(sGdTerme, 3) <> "EST" And Left$(sGdTerme, 4) <> "SONT" And _
            (Left$(sGdTerme, 3) <> "TOU" Or Not bInversion) Then
            sGdTerme = "EST COMME " & Trim$(sGdTerme)
            GoTo SuiteConclusion
        End If
    End If
    GoTo SuiteConclusion

Pos3:
    sGdTerme = Left$(glb.aBC(iNumAss1).sEntree, glb.aBC(iNumAss1).iPosFinMot2)
    If bVersionModifiee And Right$(sGdTerme, 1) = "(" Then
        bFctGdTerme = True
    End If

SuiteConclusion:
    If iPosMT_PMin + iPosMT_PMaj = 2 Then
        If Left$(sPetitTerme, 4) = "EST " Then
            sPetitTerme = Right$(sPetitTerme, Len(sPetitTerme) - 4)
            If Left$(sPetitTerme, 3) = "UN " Then
                sPetitTerme = Right$(sPetitTerme, Len(sPetitTerme) - 3)
            End If
            sPetitTerme = "QUELQUE " + Trim$(sPetitTerme) + " "
            bInversion = False
        ElseIf Left$(sPetitTerme, 5) = "SONT " Then
            sPetitTerme = Right$(sPetitTerme, Len(sPetitTerme) - 5)
            sPetitTerme = "QUELQUES " + Trim$(sPetitTerme) '+ " "
            bInversion = False
        End If
    End If

    Dim sConclusion$, sConclusionParlee$
    sPetitTerme = Trim$(sPetitTerme)
    sGdTerme = Trim$(sGdTerme)
    
    If bVersionModifiee Then ' 09/04/2017
        If bInversion Then
            If Left$(sGdTerme, 1) = "=" Then sGdTerme = Mid$(sGdTerme, 2)
        Else
            If Left$(sPetitTerme, 1) = "=" Then sPetitTerme = Mid$(sPetitTerme, 2)
        End If
    End If
    
    sConclusion = "DONC " & sPetitTerme & " " & sGdTerme
    If bInversion Then
        sConclusion = "DONC " & sGdTerme & " " & sPetitTerme
    End If
    If bFctGdTerme Then
        If bInversion Then sPetitTerme = Right$(sPetitTerme, Len(sPetitTerme) - 4)
        sConclusion = "DONC " & sGdTerme & " " & sPetitTerme & " )"
    End If
    sConclusionParlee = sConclusion
    
    If bVersionModifiee Then _
        sConclusion = sConclusion & "  <" & sMotPivot & ">"
    
    ReponseIAVB sConclusion, sTexteParleSpecifique:=sConclusionParlee
    
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
    
AssertionPrecedente:
Next iNumAss1

    If bRechecheApprofondie And Not bPasse2 Then bPasse2 = True: GoTo Passe2
    
    If bAucunMotTrouve Then _
        ReponseIAVB "Je ne peux rien conclure !", bAffiche2Etoiles:=True

End Sub

Private Sub InitMotsIgnores()

    glb.iNbAssertions = 0
    
    'Data EST, LE, LA, DE, UN, UNE
    glb.asMotsIgnores(1) = "EST"
    glb.asMotsIgnores(2) = "LE"
    glb.asMotsIgnores(3) = "LA"
    glb.asMotsIgnores(4) = "DE"
    glb.asMotsIgnores(5) = "UN"
    glb.asMotsIgnores(6) = "UNE"
    
    'Data L, DU, D, LES, DES, ET
    glb.asMotsIgnores(7) = "OR"
    glb.asMotsIgnores(8) = "L"
    glb.asMotsIgnores(9) = "DU"
    glb.asMotsIgnores(10) = "D"
    glb.asMotsIgnores(11) = "LES"
    glb.asMotsIgnores(12) = "DES"
    glb.asMotsIgnores(13) = "ET"
    
    'Data QU, QUE, QUI, SONT
    glb.asMotsIgnores(14) = "QU"
    glb.asMotsIgnores(15) = "QUE"
    glb.asMotsIgnores(16) = "QUI"
    glb.asMotsIgnores(17) = "SONT"
    
    'Data IL, ELLE, A, T, ETE
    glb.asMotsIgnores(18) = "IL"
    glb.asMotsIgnores(19) = "ELLE"
    
    glb.asMotsIgnores(20) = "A"
    glb.asMotsIgnores(21) = "T"
    glb.asMotsIgnores(22) = "ETE"
    
    'Data EN, OU, COMMENT, AU
    glb.asMotsIgnores(23) = "EN"
    glb.asMotsIgnores(24) = "OU"
    glb.asMotsIgnores(25) = "COMMENT"
    glb.asMotsIgnores(26) = "AU"
    
    'Data N, NE, S, SE, ETAIT
    glb.asMotsIgnores(27) = "N"
    glb.asMotsIgnores(28) = "NE"
    glb.asMotsIgnores(29) = "S"
    glb.asMotsIgnores(30) = "SE"
    glb.asMotsIgnores(31) = "ETAIT"
    
    'Data QUOI, C, CE, QUEL, QUELLE
    glb.asMotsIgnores(32) = "QUOI"
    glb.asMotsIgnores(33) = "C"
    glb.asMotsIgnores(34) = "CE"
    glb.asMotsIgnores(35) = "QUEL"
    glb.asMotsIgnores(36) = "QUELLE"
    
    'DATA QUELS, QUELLES, PAR,
    glb.asMotsIgnores(37) = "QUELS"
    glb.asMotsIgnores(38) = "QUELLES"
    glb.asMotsIgnores(39) = "PAR"

    'Data LEQUEL, LAQUELLE
    glb.asMotsIgnores(40) = "LEQUEL"
    glb.asMotsIgnores(41) = "LAQUELLE"
    
    'DATA CA, SIGNIFIE, TOUT, OR
    glb.asMotsIgnores(42) = "CA"
    'glb.asMotsIgnores(42) = "ÇA"
    glb.asMotsIgnores(43) = "SIGNIFIE"
    glb.asMotsIgnores(44) = "TOUT"
    'glb.asMotsIgnores(45) = "OR"
    
    'Data TOUS, TOUTE, TOUTES
    glb.asMotsIgnores(45) = "TOUS"
    glb.asMotsIgnores(46) = "TOUTE"
    glb.asMotsIgnores(47) = "TOUTES"
    
    'DATA =, (, ), Y, FF
    glb.asMotsIgnores(48) = "("
    glb.asMotsIgnores(49) = ")"
    glb.asMotsIgnores(50) = "="
    glb.asMotsIgnores(51) = "Y"
    
    'Const iNbMotsIgnores% = 51

End Sub

Private Sub InitAssertionsExemples()

    ListAssert.AddItem ("/EFF ' Effacement de toute la base")
    ListAssert.AddItem ("/D   ' Effacement de la denière assertion")
    ListAssert.AddItem ("/D1  ' Effacement de l'assertion n°1")
    ListAssert.AddItem ("/L   ' Liste des assertions de la base")
    ListAssert.AddItem ("/C   ' Copie du résultat d'exécution dans le presse papier")
    ListAssert.AddItem ("/S   ' Mettre en sourdine")
    ListAssert.AddItem ("/PARLE ' Ré-activer la voix")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("TOUT STYLO EST BLEU")
    ListAssert.AddItem ("FRANCOIS POSSEDE UN STYLO")
    ListAssert.AddItem ("BLEU EST UNE COULEUR")
    ListAssert.AddItem ("ROUGE EST UNE COULEUR")
    ListAssert.AddItem ("TOUT STYLO EST EN PLASTIQUE")
    ListAssert.AddItem ("LE PLASTIQUE EST UNE MATIERE")
    ListAssert.AddItem ("FRANCOIS POSSEDE-T-IL UN STYLO ROUGE ?")
    ListAssert.AddItem ("QUELLE EST LA COULEUR DU STYLO DE FRANCOIS ?")
    ListAssert.AddItem ("RAOUL A ACHETE UN STYLO EGALEMENT")
    ListAssert.AddItem ("DE QUELLE COULEUR EST LE STYLO DE RAOUL ?")
    ListAssert.AddItem ("TOUT STYLO EST EN QUELLE MATIERE ?")
    ListAssert.AddItem ("EN QUELLE MATIERE EST TOUT STYLO ?")
    ListAssert.AddItem ("' Limitation actuelle du logiciel : réponse fausse :")
    ListAssert.AddItem ("FRANCOIS POSSEDE-T-IL UN STYLO BLEU ?")
    ListAssert.AddItem ("FRANCOIS POSSEDE-T-IL UN STYLO ?")
    ListAssert.AddItem ("COULEUR STYLO FRANCOIS ?")
    ListAssert.AddItem ("DE QUELLE COULEUR EST LE STYLO DE FRANCOIS ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("ANNIE EST UNE JOLIE FILLE")
    ListAssert.AddItem ("ANNIE EST SAGE")
    ListAssert.AddItem ("MINNA EST UNE FILLE ELLE AUSSI")
    ListAssert.AddItem ("MINNA EST FOFOLLE")
    ListAssert.AddItem ("QUELLE FILLE EST SAGE ?")
    ListAssert.AddItem ("ET LAQUELLE EST FOFOLLE ?")
    ListAssert.AddItem ("JOLIE EST LE CONTRAIRE DE LAIDE")
    ListAssert.AddItem ("EST-CE QU'ANNIE EST LAIDE ?")
    ListAssert.AddItem ("CA SIGNIFIE-T-IL QU'ANNIE EST UNE JOLIE FILLE ?")
        
    ListAssert.AddItem ("")
    ListAssert.AddItem ("JEAN REGARDE MARIE")
    ListAssert.AddItem ("MARIE REGARDE HUGUES")
    ListAssert.AddItem ("QUI MARIE REGARDE-T-ELLE ?")
    ListAssert.AddItem ("MARIE REGARDE-T-ELLE JEAN ?")
    ListAssert.AddItem ("HUGUES EST LE FRERE D'HENRI")
    ListAssert.AddItem ("HENRI EST LE FILS D'OCTAVE")
    ListAssert.AddItem ("OCTAVE EST L'ONCLE D'ANATOLE")
    ListAssert.AddItem ("QUI EST LE FILS DE L'ONCLE D'ANATOLE ?")
    ListAssert.AddItem ("ET QUI REGARDE LE FRERE DU FILS DE L'ONCLE D'ANATOLE ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("L'ENTREPRISE A UN SIEGE ET UNE ANNEXE")
    ListAssert.AddItem ("M.BERTRAND EST RESPONSABLE DE L'ANNEXE")
    ListAssert.AddItem ("L'ANNEXE A 15 SERVICES DIFFERENTS")
    ListAssert.AddItem ("M.JACQUES EST RESPONSABLE DE LA SAISIE")
    ListAssert.AddItem ("M.RENE EST LE CHEF DU SERVICE COMPTABILITE")
    ListAssert.AddItem ("M.MARTIN EST UN AMI DE M.JACQUES")
    ListAssert.AddItem ("M.DUBOIS EST CHEF DU SERVICE PHOTO")
    ListAssert.AddItem ("LA SAISIE EST UN SERVICE DECENTRALISE")
    ListAssert.AddItem ("RENAUD EST LE FILS DE M.BERTRAND")
    ListAssert.AddItem ("DAMIEN EST LE FILS DE M.RENE")
    ListAssert.AddItem ("CHEF DE SERVICE ?")
    ListAssert.AddItem ("CHEF DE SERVICE PHOTO ?")
    ListAssert.AddItem ("SERVICE PHOTO ?")
    ListAssert.AddItem ("RESPONSABLE ?")
    ListAssert.AddItem ("M.MARTIN EST L'AMI DE QUI ?")
    ListAssert.AddItem ("QUI EST LE FILS DU RESPONSABLE DES 15 SERVICES ?")
    ListAssert.AddItem ("QUI EST LE FILS D'UN CHEF DE SERVICE ?")
    ListAssert.AddItem ("QUI EST L'AMI DU RESPONSABLE D'UN SERVICE DECENTRALISE ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("TOUT HOMME EST MORTEL")
    ListAssert.AddItem ("OR SOCRATE EST UN HOMME")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("TOUT HOMME EST BIPEDE")
    ListAssert.AddItem ("OR PAUL EST UN HOMME")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("TOUT CE QUI EST RARE EST CHER")
    'ListAssert.AddItem ("UN CHEVAL BON_MARCHE EST RARE")
    ListAssert.AddItem ("UN CHEVAL BON.MARCHE EST RARE")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("PLATON EST GREC")
    ListAssert.AddItem ("OR PLATON EST PHILOSOPHE")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("QUEL PHILOSOPHE EST GREC ?")
    ListAssert.AddItem ("Y A-T-IL UN PHILOSOPHE GREC ?")
    ListAssert.AddItem ("QUI EST GREC ET PHILOSOPHE ?")
    ListAssert.AddItem ("PHILOSOPHE GREC ?")
    ListAssert.AddItem ("PLATON EST-IL GREC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("LES SAVANTS SONT SOUVENT DISTRAITS")
    ListAssert.AddItem ("OR TOUS LES SAVANTS SONT BAVARDS")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("TOUS LES HOMMES SONT DES MORTELS")
    ListAssert.AddItem ("OR DES HOMMES SONT JUSTES")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("X1 = X2")
    ListAssert.AddItem ("X2 = X3")
    ListAssert.AddItem ("PAR QUOI X1 = X3 ?")
    
    ListAssert.AddItem ("")
    'ListAssert.AddItem ("Y1 = F(X)")
    'ListAssert.AddItem ("F(X) ?")
    'ListAssert.AddItem ("X = G(T1)")
    'ListAssert.AddItem ("V = K(W)")
    'ListAssert.AddItem ("T1 = H(V)")
    'ListAssert.AddItem ("F(G(H(K(W)))) ?")
    ListAssert.AddItem ("Y1 = F DE X")
    ListAssert.AddItem ("F DE X ?")
    ListAssert.AddItem ("X = G DE T1")
    ListAssert.AddItem ("V = K DE W")
    ListAssert.AddItem ("T1 = H DE V")
    ListAssert.AddItem ("F DE G DE H DE K DE W ?")
    ListAssert.AddItem ("A QUOI EST = F DE G DE H DE K DE W ?")
    ListAssert.AddItem ("F G H K W ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("CHAT = ANIMAL")
    ListAssert.AddItem ("CHAT = MANGEUR(SOURIS)")
    'ListAssert.AddItem ("CHAT = MANGEUR DE SOURIS")
    ListAssert.AddItem ("TIGRE = ANIMAL")
    ListAssert.AddItem ("TIGRE = MANGEUR(HOMME)")
    ' Problème de syntaxe des syllogismes basée sur la détection des ()
    '  si on change, certaines déductions ne marche plus :
    'ListAssert.AddItem ("TIGRE = MANGEUR D'HOMME")
    ListAssert.AddItem ("ANIMAL MANGEUR ?")
    'ListAssert.AddItem ("MANGEUR(HOMME) = ?")
    ListAssert.AddItem ("MANGEUR D'HOMME = ?")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("DONC ?")
    'ListAssert.AddItem ("Y A-T-IL UN MANGEUR(SOURIS ET HOMME) ?")
    ListAssert.AddItem ("Y A-T-IL UN MANGEUR DE SOURIS ET D'HOMME ?")
    'ListAssert.AddItem ("QUI EST MANGEUR(HOMME ET SOURIS) ?")
    ListAssert.AddItem ("QUI EST MANGEUR D'HOMME ET DE SOURIS ?")
    
    ListAssert.AddItem ("")
    'ListAssert.AddItem ("ARTABAN = CHEVAL(HENRI_IV)")
    ListAssert.AddItem ("ARTABAN = CHEVAL DE HENRI.4")
    'ListAssert.AddItem ("BLANC = COULEUR(ARTABAN)")
    ListAssert.AddItem ("BLANC = COULEUR D'ARTABAN")
    'ListAssert.AddItem ("HENRI_IV = ROI(NAVARRE)")
    ListAssert.AddItem ("HENRI.4 = ROI DE NAVARRE")
    'ListAssert.AddItem ("LOUIS_14 = ROI(FRANCE)")
    ListAssert.AddItem ("LOUIS.14 = ROI DE FRANCE")
    'ListAssert.AddItem ("COULEUR(CHEVAL(ROI(NAVARRE))) ?")
    ListAssert.AddItem ("COULEUR DU CHEVAL DU ROI DE NAVARRE ?")
    ListAssert.AddItem ("QUELLE EST LA COULEUR DU CHEVAL DU ROI DE NAVARRE ?")
    ListAssert.AddItem ("QUELLE EST LA COULEUR DU CHEVAL BLANC DU ROI DE NAVARRE ?")
    ListAssert.AddItem ("' Faille : composition incomplète, mauvaise indirection :")
    ListAssert.AddItem ("QUELLE EST LA COULEUR DU CHEVAL DU ROI ?")
    ListAssert.AddItem ("' Bonne indirection :")
    ListAssert.AddItem ("QUEL ROI A UN CHEVAL ?")
    ListAssert.AddItem ("' Faille :  mauvaise indirection :")
    ListAssert.AddItem ("CHEVAL DU ROI ?")
    ListAssert.AddItem ("CHEVAL DU ROI DE NAVARRE ?")
    ListAssert.AddItem ("CHEVAL DU ROI DE FRANCE ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("LE CANARI EST UN OISEAU JAUNE")
    ListAssert.AddItem ("JAUNE EST UNE COULEUR")
    ListAssert.AddItem ("QUEL OISEAU EST JAUNE ?")
    ListAssert.AddItem ("QUEL EST L'OISEAU JAUNE ?")
    ListAssert.AddItem ("DE QUELLE COULEUR EST LE CANARI ?")
    ListAssert.AddItem ("QUELLE EST LA COULEUR DU CANARI ?")
    ListAssert.AddItem ("COULEUR CANARI ?")
    ListAssert.AddItem ("COULEUR ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("MARSEILLE EST LA VILLE PHOCEENNE")
    ListAssert.AddItem ("DEFERRE EST LE MAIRE DE MARSEILLE")
    ListAssert.AddItem ("PHOCEENNE SIGNIFIE ORIGINAIRE DE PHOCEE")
    ListAssert.AddItem ("LA PHOCEE EST UNE PROVINCE GRECQUE")
    ListAssert.AddItem ("GASTON EST LE PRENOM DE DEFERRE")
    ListAssert.AddItem ("QUEL EST LE PRENOM DU MAIRE DE MARSEILLE ?")
    ListAssert.AddItem ("QUEL EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE GRECQUE ?")
    ListAssert.AddItem ("' Test du système : composition de fonctions incomplète :")
    ListAssert.AddItem ("QUEL EST LE PRENOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE ?")
    ListAssert.AddItem ("' Autres réponses souhaitées : GASTON EST LE PRÉNOM DU MAIRE DE LA VILLE ORIGINAIRE D'UNE PROVINCE PHOCÉENNE")
    ListAssert.AddItem ("' ou bien : D'UNE PROVINCE PHOCÉENNE ?")
    ListAssert.AddItem ("' ou bien : ** heu... GASTON ?")
            
    ListAssert.AddItem ("")
    ListAssert.AddItem ("PAUL POSSEDE UN PERROQUET BAVARD")
    ListAssert.AddItem ("MULTICOLORE SIGNIFIE DE PLUSIEURS COULEURS")
    ListAssert.AddItem ("UN PERROQUET EST UN ANIMAL MULTICOLORE")
    ListAssert.AddItem ("QUI POSSEDE UN ANIMAL DE PLUSIEURS COULEURS ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("MARIE EST UNE JOLIE FILLE")
    ListAssert.AddItem ("JOLIE EST LE CONTRAIRE DE LAIDE")
    ListAssert.AddItem ("MARIE EST-ELLE LAIDE ?")
    ListAssert.AddItem ("EST-CE QUE MARIE EST LAIDE ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("SABINE AIME JACQUES")
    ListAssert.AddItem ("QUI AIME JACQUES ?")
    ListAssert.AddItem ("QUI AIME SABINE ?")
    ListAssert.AddItem ("QUI JACQUES AIME-T-IL ?")
    ListAssert.AddItem ("' Meilleure réponse : Je L'ignore :")
    ListAssert.AddItem ("JACQUES AIME-T-IL SABINE ?")
    ListAssert.AddItem ("SABINE AIME-T-ELLE JACQUES ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("' Capacités de la version modifiée :")
    ListAssert.AddItem ("' Ordre des termes : DONC TOUT CHEVAL EST HERBIVORE :")
    ListAssert.AddItem ("TOUT CHEVAL EST UN EQUIDE")
    ListAssert.AddItem ("OR TOUT EQUIDE EST HERBIVORE")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("' Sens logique : DONC PAUL EST COMME TOUT HOMME :")
    ListAssert.AddItem ("TOUT HOMME EST BIPEDE")
    ListAssert.AddItem ("OR PAUL EST BIPEDE")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("' Sens logique : ARISTOTE EST COMME TOUT HOMME :")
    ListAssert.AddItem ("TOUT HOMME EST RATIONNEL")
    ListAssert.AddItem ("OR ARISTOTE EST RATIONNEL")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("ARISTOTE ETAIT GREC")
    ListAssert.AddItem ("ARISTOTE ETAIT PHILOSOPHE")
    ListAssert.AddItem ("PHILOSOPHE GREC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("' Syllogismes : ordre des assertions")
    ListAssert.AddItem ("OR SOCRATE EST UN HOMME")
    ListAssert.AddItem ("TOUT HOMME EST MORTEL")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("' Syllogismes : plusieurs à la fois")
    ListAssert.AddItem ("TOUT HOMME EST MORTEL")
    ListAssert.AddItem ("TOUT HOMME EST BIPEDE")
    ListAssert.AddItem ("OR SOCRATE EST UN HOMME")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("OR PAUL EST UN HOMME")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("TOUT LOGICIEN EST INCOMPRIS")
    ListAssert.AddItem ("OR TOUT HOMME SENSE EST LOGICIEN")
    ListAssert.AddItem ("DONC ?")
    
    ListAssert.AddItem ("")
    ListAssert.AddItem ("' Syllogismes : syntaxe imparfaite (de la version modifiée)")
    ListAssert.AddItem ("LE LOUVRE EST BEAU")
    ListAssert.AddItem ("OR J'AIME TOUT CE QUI EST BEAU")
    ListAssert.AddItem ("DONC ?")
    ListAssert.AddItem ("")
    
End Sub

Private Sub CmdSuivant_Click()
    On Error Resume Next
    If ListAssert.ListIndex < iNumLigneDebutExemples Then
        ListAssert.ListIndex = iNumLigneDebutExemples
    Else
        ListAssert.ListIndex = ListAssert.ListIndex + 1
    End If
    TraiterAssertion
End Sub

Private Sub Initialiser()
    
    If bInitAgent() Then
        glb.bAgent = True
        glb.bVoixActive = True
    End If
    
    InitMotsIgnores
    InitAssertionsExemples
    
End Sub

#If bVersionVBA Then
Private Function bInitAgent() As Boolean
    bInitAgent = False
End Function
#Else ' Version VB6
#End If

Private Sub ListAssert_Click()
    TextInput.Text = ListAssert
End Sub

#If bVersionVBA Then
Private Sub ListAssert_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    TraiterAssertion
End Sub
#Else ' Version VB6
Private Sub ListAssert_DblClick()
    TraiterAssertion
End Sub
#End If

Private Sub TraiterAssertion()
    'TextInput.Text = ListAssert.Text
    'TextInput.Text = ListAssert
    TextInput.Text = ListAssert.List(ListAssert.ListIndex)
    CmdGo_Click
End Sub

Private Sub CmdGo_Click()
    
    Dim bMemVoixActive As Boolean
    bMemVoixActive = glb.bVoixActive
    
    If glb.bVoixActive Then
        ' Le contrôle ocx TextToSpeec à l'air de sévèrement bugger :
        '  limiter la réentrance dans les fonctions
        CmdGo.Enabled = False
        CmdSuivant.Enabled = False
        TextInput.Enabled = False
        ListAssert.Enabled = False
        ListIA.Enabled = False
    End If
    
    IAVBMain
    PositionnerListIA
    
    If bMemVoixActive Then
        CmdGo.Enabled = True
        CmdSuivant.Enabled = True
        CmdSuivant.SetFocus
        TextInput.Enabled = True
        ListAssert.Enabled = True
        ListIA.Enabled = True
    End If
    
End Sub

Private Sub PositionnerListIA()
    ' Toujours se positionner sur la dernière ligne
    bPositionnement = True
    ListIA.ListIndex = ListIA.ListCount - 1
    bPositionnement = False
End Sub

Private Sub RecopierLigneListIA()
    If Not bPositionnement Then TextInput.Text = ListIA
End Sub

Private Sub ListIA_Click()
    RecopierLigneListIA
End Sub

#If bVersionVBA Then
Private Sub ListIA__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RecopierLigneListIA
    CmdGo_Click
End Sub
#Else ' Version VB6
Private Sub ListIA_DblClick()
    RecopierLigneListIA
    CmdGo_Click
End Sub
#End If

#If bVersionVBA Then
Private Sub TextInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = iCodeToucheEntree Then
        IAVBMain
        PositionnerListIA
    End If
End Sub
#Else ' Version VB6
Private Sub TextInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = iCodeToucheEntree Then ' Touche Entrée
        IAVBMain
        PositionnerListIA
    End If
End Sub
#End If

#If bVersionVBA Then
Private Sub TextInput_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
#Else ' Version VB6
Private Sub TextInput_KeyUp(KeyCode As Integer, Shift As Integer)
#End If
    ' Enlever la touche entrée pour la zone multiligne
    Dim I%, iLen%
    iLen = Len(TextInput)
    For I = 1 To iLen
        If Len(Mid$(TextInput, I, 1)) > 0 Then
            If Asc(Mid$(TextInput, I, 1)) = iCodeToucheEntree Then
                TextInput = Mid$(TextInput, 1, I - 1) & Mid$(TextInput, I + 2, iLen - I - 1)
            End If
        End If
    Next I
End Sub

