
' Fichier frmIAVB.vb
' ------------------

Imports System.Collections.Generic
Imports System.Speech
Imports System.Speech.Synthesis

Namespace IAVB

Friend Class frmIAVB : Inherits Form

#Region "Configuration"

' Booléen pour forcer le fonctionnement dans la version originale (1984)
'  ou bien modifiée (2001)
Private Const bVersionModifiee As Boolean = True ' True par défaut

' Normaliser les espaces et sauts de ligne en sortie
Private Const bNormalisationSortieTrimEtVbLf As Boolean = True ' True par défaut

' Convertir tout en minuscules (pour éviter d'avoir à taper des majuscules avec accent,
'  et pour pouvoir prononcer "ça" via MS-Agent)
Private Const bTraiterEnMinuscules As Boolean = True ' True par défaut

' 29/04/2017 Enlever les accents (pour pouvoir comparer avec la version originale de 1984)
Private Const bTraiterSansAccents As Boolean = False ' False par défaut

#End Region

#Region "Initialisations"

Private Const iCodeToucheEntree% = 13
Private Const iNumLigneDebutExemples% = 8 ' Pour la démo sur les exemples

Private m_bMSAgent As Boolean

' Pour pouvoir recopier la ligne ListIA dans la fonction RecopierLigneListIA
Private m_bPositionnement As Boolean

Private m_oIAVB As New IAVB.clsIAVB

' SpeechSynthesizer Class Provides access to the functionality of an installed a speech synthesis engine.
Private WithEvents m_synth As New SpeechSynthesizer
Const sVoixFR$ = "Microsoft Hortense Desktop"
Const sVoix_enUSWin10$ = "Microsoft Zira Desktop"
Const sVoix_enUSWin7$ = "Microsoft Anna"
Const sFrancais$ = "Français (Hortense)"
Const sAnglais$ = "Anglais (Zira)"
Private m_sVoix$ = ""

Const sModeVocalDef$ = sFrancais
Const sModeSilencieux$ = "Silencieux"
Const sModeMSAgent$ = "MS-Agent"
Const iIndexModeSilencieux% = 0
'Const iIndexModeMSAgent% = 1
Dim iIndexModeSynthVocWin% = 2

Const sTxtCmdSuiv$ = "->"
Const sMsgCmdSuiv$ = "Exemple suivant"
Const sTxtCmdSuivStop$ = "St."
Const sMsgCmdSuivStop$ = "Stop"

Private Sub Initialisations()

    Dim sTxt$ = Me.Text & " - V" & sVersionAppli & " (" & sDateVersionAppli & ")"
    If bDebug Then sTxt &= " - Debug"
    Dim b64 As Boolean = Is64BitProcess()
    If b64 Then sTxt &= " - 64 bits" Else sTxt &= " - 32 bits"
    Me.Text = sTxt

    Dim bVoixActive As Boolean
    If Not b64 Then VerifierInstallAgent(m_bMSAgent, bVoixActive)

    iIndexModeSynthVocWin = 1
    Me.listParole.Items.Add(sModeSilencieux)
    If m_bMSAgent Then Me.listParole.Items.Add(sModeMSAgent) : iIndexModeSynthVocWin = 2
    InitVoix()
    Me.listParole.SelectedIndex = iIndexModeSilencieux

    m_oIAVB.Initialiser(bVersionModifiee, bNormalisationSortieTrimEtVbLf, _
        bTraiterEnMinuscules, bTraiterSansAccents)
    InitAssertionsExemples()

End Sub

Private Sub InitVoix()

    ' SpeechSynthesizer.GetInstalledVoices, méthode (CultureInfo)
    ' https://msdn.microsoft.com/fr-fr/library/ms586870(v=vs.110).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
    ' Output information about all of the installed voices that
    '  support the en-US locacale. 
    Dim dicoVoix As New Dictionary(Of String, VoiceInfo)
    For Each voix As InstalledVoice In m_synth.GetInstalledVoices() 'New CultureInfo("en-US"))
        Dim info As VoiceInfo = voix.VoiceInfo
        Dim sCle$ = info.Name & " (" & info.Culture.Name & ", " & info.Gender.ToString() & ")"
        ' Remplacer les noms des deux voix standards de Windows 10 par Français et Anglais
        If info.Name = sVoixFR Then sCle = sFrancais
        If info.Name = sVoix_enUSWin10 Then sCle = sAnglais
        If info.Name = sVoix_enUSWin7 Then sCle = sAnglais
        If Not dicoVoix.ContainsKey(sCle) Then
            dicoVoix.Add(sCle, info)
            Me.listParole.Items.Add(sCle)
        End If
    Next

End Sub

Private Sub listParole_SelectedIndexChanged(sender As Object, e As EventArgs) _
    Handles listParole.SelectedIndexChanged

    If listParole.Text = sModeMSAgent Then
        If Not bInitMSAgent(m_bMSAgent, m_oIAVB.m_bVoixActive) Then
            listParole.Text = sModeSilencieux
        End If
    Else
        bMasquerMSAgent()
    End If

    If listParole.Text = sModeSilencieux Then
        m_oIAVB.m_bVoixActive = False
    Else
        m_oIAVB.m_bVoixActive = True
    End If

    If listParole.SelectedIndex >= iIndexModeSynthVocWin Then
        Dim sLangue$ = listParole.Text
        Dim sVoix$ = ""
        If sLangue = sFrancais Then
            sVoix = sVoixFR
        ElseIf sLangue = sAnglais Then
            sVoix = sVoix_enUSWin10
        Else
            sVoix = sLangue
            Dim iPos% = sVoix.IndexOf("(")
            If iPos > -1 Then sVoix = sVoix.Substring(0, iPos - 1).Trim
        End If
        m_sVoix = sVoix
    End If

End Sub

Public Function bDire(sParole$) As Boolean

    If Me.listParole.Text = sModeMSAgent Then
        ' Les minuscules améliorent la prononciation des MS-Agent
        If m_bMSAgent AndAlso bTraiterEnMinuscules Then sParole = sParole.ToLower
        Return bDireMSAgent(sParole)
    End If

    If listParole.Text = sModeSilencieux Then Return False

    If Not IsNothing(m_synth) AndAlso m_sVoix.Length > 0 Then

        While m_synth.State = SynthesizerState.Speaking
            Attendre()
        End While

        If m_synth.State = SynthesizerState.Ready Then
            m_synth.SelectVoice(m_sVoix)
            Activation(bDesactiver:=True)
            'm_synth.SpeakAsync(sParole)
            m_synth.Speak(sParole) ' Attendre la fin
            Activation()
        End If
    Else
        listParole.Text = sModeSilencieux
    End If

    Return True

End Function

' Si SpeakAsync
'Private Sub MyEventHandler( eventSender As Object, _
'    eventArgs As SpeakCompletedEventArgs) Handles m_synth.SpeakCompleted
'    Activation()
'End Sub

Private Sub Activation(Optional bDesactiver As Boolean = False)

    ' Non, car on ne peut plus réactiver Me.chkAuto ou faire Stop
    'Me.Enabled = bDesactiver

    Me.chkAuto.Enabled = True ' Tjrs actif
    Me.cmdSuivant.Enabled = True ' Tjrs actif (Stop le mode Auto)

    Me.listAssert.Enabled = Not bDesactiver
    Me.listIA.Enabled = Not bDesactiver
    Me.textInput.Enabled = Not bDesactiver
    Me.cmdGo.Enabled = Not bDesactiver
    Me.cmdInstall.Enabled = Not bDesactiver

End Sub

Private Sub AjouterExemple(sExemple$)
    Me.listAssert.Items.Add(sExemple)
End Sub

Private Sub InitAssertionsExemples()

    AjouterExemple(m_oIAVB.sCmdEffBase & " ' Effacer toute la base")
    AjouterExemple(m_oIAVB.sCmdEff & " ' Effacer la denière assertion")
    AjouterExemple(m_oIAVB.sCmdEff & "1 ' Effacer l'assertion n°1")
    AjouterExemple(m_oIAVB.sCmdLister & " ' Lister les assertions de la base")
    AjouterExemple(m_oIAVB.sCmdCopier & " ' Copier la discussion dans le presse papier")
    AjouterExemple(m_oIAVB.sCmdSilence & " ' Mettre en sourdine (autre poss.: masquer l'agent)")
    AjouterExemple(m_oIAVB.sCmdParler & " ' Ré-activer la voix")

    Dim lst As New List(Of String)
    Dim ex As New IAVB.clsInitIAVB(lst)
    ex.InitAssertionsExemples()
    For Each sExemple As String In lst
        AjouterExemple(sExemple)
    Next

End Sub

#End Region

Private Sub TraiterCmd()

    Dim bMemVoixActive As Boolean = m_oIAVB.m_bVoixActive

    ' Eviter la réentrance dans les fonctions
    If m_oIAVB.m_bVoixActive Then Activation(bDesactiver:=True)

    ' Si la zone de saisie est multiligne, faire un Trim
    Dim sTxt$ = textInput.Text.Trim()
    m_oIAVB.IAVBMain(sTxt)

    If m_oIAVB.m_bCopierPressePapier Then
        If bCopierPressePapier(m_oIAVB.m_sDiscussion) Then
            m_oIAVB.CopiePressePapierOk()
        Else
            m_oIAVB.CopiePressePapierEchec()
        End If
    End If

    If Me.textInput.Text.Length = 0 Then
        Me.listIA.Items.Add("")
        PositionnerListIA()
    End If

    ' Activer ou réactiver la parole
    If Not bMemVoixActive AndAlso m_oIAVB.m_bVoixActive Then listParole.Text = sModeVocalDef

    Dim sReponseVocale$ = m_oIAVB.m_sReponseVocale
    If sReponseVocale.Length > 0 Then

        Dim asReponses$() = sReponseVocale.Split(vbLf.ToCharArray)
        Dim i% = 0
        For Each sReponse As String In asReponses
            If sReponse.Trim.Length = 0 Then Continue For
            ' 12/05/2017
            If i = 0 AndAlso asReponses.Length >= 1 AndAlso _
               asReponses(1) = asReponses(0) Then
                ' Rappel de la question : dans ce cas on peut afficher le texte en 1er

                ' Réponse texte
                Me.listIA.Items.Add(sReponse)
                PositionnerListIA()

                bDire(sReponse) ' Réponse vocale

            ElseIf i = 1 AndAlso asReponses.Length >= 1 AndAlso _
               asReponses(1) = asReponses(0) Then
                ' Déjà traité
            Else
                If i Mod 2 = 0 Then
                    bDire(sReponse) ' Réponse vocale
                Else
                    ' Réponse texte
                    Me.listIA.Items.Add(sReponse)
                    PositionnerListIA()
                End If
            End If
            i += 1
        Next

        ' Désactiver la parole
        If Not m_oIAVB.m_bVoixActive Then listParole.Text = sModeSilencieux

        GoTo Fin

    End If

    If m_oIAVB.m_sRappelQuestion.Length > 0 Then
        Me.listIA.Items.Add(m_oIAVB.m_sRappelQuestion)
        PositionnerListIA()
    End If

    If m_oIAVB.m_sReponse.Length > 0 Then
        Dim asReponses$() = m_oIAVB.m_sReponse.Split(vbLf.ToCharArray)
        For Each sReponse As String In asReponses
            If sReponse.Trim.Length = 0 Then Continue For
            Me.listIA.Items.Add(sReponse)
        Next
        PositionnerListIA()
    End If

Fin:
    If bMemVoixActive Then
        Activation()
        cmdSuivant.Focus()
    End If

End Sub

Private Sub RecopierLigneListIA()
    If Not m_bPositionnement Then textInput.Text = listIA.Text
End Sub

Public Sub PositionnerListIA()
    ' Toujours se positionner sur la dernière ligne
    m_bPositionnement = True
    listIA.SelectedIndex = listIA.Items.Count - 1
    m_bPositionnement = False
End Sub

Private Sub TraiterAssertion()

    textInput.Text = listAssert.Text

    ' Veiller à ce que la question soit bien affichée avant d'activer la synthèse vocale
    TraiterMsgSysteme_DoEvents()

    TraiterCmd()

End Sub

Private Sub AssertionSuivante()
    If listAssert.SelectedIndex < iNumLigneDebutExemples Then
        listAssert.SelectedIndex = iNumLigneDebutExemples
    ElseIf listAssert.SelectedIndex < listAssert.Items.Count - 1 Then
        listAssert.SelectedIndex = listAssert.SelectedIndex + 1
    End If
    TraiterAssertion()
End Sub

#Region "Gestion des événements"

Private Sub frmIAVB_Load(eventSender As Object, eventArgs As EventArgs) Handles MyBase.Load
    Initialisations()
End Sub

Private Sub cmdGo_Click(eventSender As Object, eventArgs As EventArgs) Handles cmdGo.Click
    TraiterCmd()
End Sub

Private Sub chkAuto_Click(sender As Object, e As EventArgs) Handles chkAuto.Click
    If Me.chkAuto.Checked Then
        'Me.cmdSuivant.Text = sTxtCmdSuivStop
        'Me.ToolTip1.SetToolTip(Me.cmdSuivant, sMsgCmdSuivStop)
    Else
        Me.cmdSuivant.Text = sTxtCmdSuiv
        Me.ToolTip1.SetToolTip(Me.cmdSuivant, sMsgCmdSuiv)
    End If
End Sub

Private Sub cmdSuivant_Click(eventSender As Object, eventArgs As EventArgs) _
    Handles cmdSuivant.Click

    Dim sMemAssertion2$, sMemAssertion$
    sMemAssertion = ""
    Static bEnCours As Boolean = False
    If bEnCours Then GoTo Fin

Recommencer:
    bEnCours = True
    sMemAssertion2 = sMemAssertion
    sMemAssertion = Me.listIA.Text.Trim
    AssertionSuivante()

    If Me.chkAuto.Checked Then

        Me.cmdSuivant.Text = sTxtCmdSuivStop
        Me.ToolTip1.SetToolTip(Me.cmdSuivant, sMsgCmdSuivStop)

        If Me.listIA.Text.Trim.Length = 0 AndAlso _
           sMemAssertion.Length = 0 AndAlso sMemAssertion2.Length = 0 Then
            Me.textInput.Text = m_oIAVB.sCmdCopier ' "/C"
            TraiterCmd()
            bEnCours = False
            Exit Sub
        End If
        ' Voir si l'utilisateur à cliquer sur Stop
        TraiterMsgSysteme_DoEvents()
        GoTo Recommencer
    End If

Fin:
    bEnCours = False
    Me.chkAuto.Checked = False
    Me.cmdSuivant.Text = sTxtCmdSuiv
    Me.ToolTip1.SetToolTip(Me.cmdSuivant, sMsgCmdSuiv)

End Sub

Private Sub ListAssert_SelectedIndexChanged(eventSender As Object, _
    eventArgs As EventArgs) Handles listAssert.SelectedIndexChanged
    textInput.Text = listAssert.Text
End Sub

Private Sub ListAssert_DoubleClick(eventSender As Object, eventArgs As EventArgs) _
    Handles listAssert.DoubleClick
    TraiterAssertion()
End Sub

Private Sub listIA_SelectedIndexChanged(eventSender As Object, eventArgs As EventArgs) _
    Handles listIA.SelectedIndexChanged
    RecopierLigneListIA()
End Sub

Private Sub listIA_DoubleClick(eventSender As Object, eventArgs As EventArgs) _
    Handles listIA.DoubleClick
    RecopierLigneListIA()
    TraiterCmd()
End Sub

Private Sub textInput_KeyPress(eventSender As Object, eventArgs As KeyPressEventArgs) _
    Handles textInput.KeyPress

    Dim iCodeTouche% = Asc(eventArgs.KeyChar)
    If iCodeTouche = iCodeToucheEntree Then ' Touche Entrée
        TraiterCmd()
        Exit Sub
    End If

End Sub

Private Sub cmdInstall_Click(sender As Object, e As EventArgs) Handles cmdInstall.Click
    bInstallationMSAgent(m_bMSAgent, m_oIAVB.m_bVoixActive)
End Sub

#End Region

End Class

End Namespace