
' Fichier modSynthVocale.vb : Module de synthèse vocale via MS-Agent
' -------------------------

Option Strict Off ' Liaison tardive ('On' possible en liaison anticipée)

Module modSynthVocaleMSAgent

#Region "Déclarations"

' Liaison anticipée : ajouter les références vers 
'  Interop.AgentObjects.dll et AxInterop.AgentObjects.dll 
'Imports AgentObjects ' Pour Agent et IAgentCtlCharacter
'Private WithEvents m_oAgent As Agent ' WithEvents pour recevoir l'év. _RequestComplete
'Private m_oMerlin As IAgentCtlCharacter
'Private m_oMerlinRequest As IAgentCtlRequest

' Liaison tardive
Private m_oAgent As Object
Private m_oMerlin As Object

Private Const sMSAgentCtrl2$ = "Microsoft Agent Control 2.0"
Private Const sMSAgentCtrl2ProgID$ = "Agent.Control.2"
'Private Const sMSAgentCtrl2Ver$ = "2 du ??/??/???? : ? Ko"
'Private Const sMSAgentCtrl2Dll$ = "agentctl.dll"
Private Const sMSAgentInst1Txt$ = "Microsoft Agent core components (MSagent.exe : 392 Ko)"
Private Const sMSAgentInst1$ = "MSagent.exe"
Private Const SMSAgentURL1$ = "http://activex.microsoft.com/activex/controls/agent2/MSagent.exe"
Private Const sMSAgentInst2Txt$ = "Language component : French (AgtX040C.exe : 129 Ko)"
Private Const sMSAgentInst2$ = "AgtX040C.exe"
Private Const sMSAgentInst3Txt$ = "Character : Merlin (Merlin.exe : 1830 Ko)"
Private Const sMSAgentInst3$ = "Merlin.exe"
Private Const SMSAgentURL$ = "http://www.microsoft.com/msagent/downloads/user.aspx"
Private Const sMSAgentACS$ = "merlin.acs"

Private Const sSpeechAPI$ = "Microsoft Speech Object Library"
'Private Const sSpeechAPIProgID$ = "SAPI.SpLexicon.1"
Private Const sSpeechAPIDll$ = "sapi.dll"
Private Const sSpeechAPIInst$ = "spchapi.exe"
Private Const sSpeechAPIInstTxt$ = "SAPI 4.0 runtime support"
Private Const sSpeechAPI_URL$ = "http://activex.microsoft.com/activex/controls/sapi/spchapi.exe"

Private Const sLHTTS3000Fr$ = "L&H TTS3000 Text-To-Speech - French"
'Private Const sLHTTS3000FrClsID$ = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}" ' ClassID et non ProgID
'Private Const sLHTTS3000FrVer$ = "3000 du ??/??/???? : ? Ko"
Private Const sLHTTS3000FrDll$ = "ttsFRFwr.dll"
Private Const sLHTTS3000FrInst$ = "lhttsfrf.exe"
Private Const sLHTTS3000FrInstTxt$ = "Lernout & Hauspie® TTS3000 TTS engine - French (2.2 MB exe)"

#End Region

#Region "Installation"

Public Sub VerifierInstallAgent(ByRef bAgent As Boolean, ByRef bVoixActive As Boolean)

    bAgent = False
    bVoixActive = False

    If Not bVerifierInstallObjet(sMSAgentCtrl2ProgID) Then Exit Sub

    Dim sWinDir$ = sDossierParent(Environment.GetFolderPath( _
        Environment.SpecialFolder.System))

    Dim sCheminMerlin$ = sWinDir & "\msagent\chars\" & sMSAgentACS
    If bFichierExiste(sCheminMerlin) Then bAgent = True

    Dim bTTS_Ok As Boolean
    Dim sCheminTTS_LH$ = sWinDir & "\Lhsp\" & sLHTTS3000FrDll
    If bFichierExiste(sCheminTTS_LH) Then bTTS_Ok = True
    ' Si on désenregistre la dll et qu'on la supprime
    '  la classe est encore dans la BR : test non discriminant :
    '  simplement vérifier la présence de la dll
    'If bVerifierInstallObjet(sLHTTS3000FrClsID, bClassID:=True) Then bTTS_Ok = True

    'Dim bSAPI_Ok As Boolean
    'Dim sFichiersCommuns$ = Environment.GetFolderPath( _
    '    Environment.SpecialFolder.CommonProgramFiles)
    'Dim sCheminSAPI$ = sFichiersCommuns & "\Microsoft Shared\Speech\" & sSpeechAPIDll
    'If bFichierExiste(sCheminSAPI) Then bSAPI_Ok = True

    'If bAgent And bTTS_Ok And bSAPI_Ok Then bVoixActive = True
    ' 01/05/2015 Plus besoin de SAPI avec Windows 7 64b ?
    If bAgent And bTTS_Ok Then bVoixActive = True

End Sub

Public Function bInstallationMSAgent( _
    ByRef bMSAgent As Boolean, ByRef bVoixActive As Boolean) As Boolean

    bMSAgent = False
    bVoixActive = False

    If Not bVerifierInstallObjet(sMSAgentCtrl2ProgID) Then

        If Not bVerifierDllActiveX_InstExe(sMSAgentCtrl2, sMSAgentCtrl2ProgID, _
            sMSAgentInst1, sMSAgentInst1Txt, SMSAgentURL1) Then GoTo MsgReinst

        If Not bVerifierDllActiveX_InstExe(sMSAgentCtrl2, sMSAgentCtrl2ProgID, _
            sMSAgentInst2, sMSAgentInst2Txt, SMSAgentURL) Then GoTo MsgReinst

        GoTo MsgReinst

    End If

    Dim sWinDir$ = sDossierParent(Environment.GetFolderPath( _
        Environment.SpecialFolder.System))

    ' Si cette dernière étape est validée, alors MSAgent complètement installé
    'C:\WINDOWS\msagent\chars\merlin.acs
    Dim sCheminMerlin$ = sWinDir & "\msagent\chars\" & sMSAgentACS
    If Not bVerifierDllActiveX_InstExe(sMSAgentCtrl2, sMSAgentCtrl2ProgID, _
        sMSAgentInst3, sMSAgentInst3Txt, SMSAgentURL, sCheminMerlin) Then _
        GoTo MsgReinst

    ' Reste la synthèse vocale (MS Agent peut être installé sans voix)

    Dim sCheminTTS_LH$ = sWinDir & "\Lhsp\" & sLHTTS3000FrDll
    If Not bVerifierDllActiveX_InstExe(sLHTTS3000Fr, "", _
        sLHTTS3000FrInst, sLHTTS3000FrInstTxt, SMSAgentURL, sCheminTTS_LH, _
        bClassID:=True) Then GoTo MsgReinst

    'C:\Program Files\Fichiers communs\Microsoft Shared\Speech\sapi.dll
    Dim sFC$ = Environment.GetFolderPath( _
        Environment.SpecialFolder.CommonProgramFiles)
    Dim sCheminSAPI$ = sFC & "\Microsoft Shared\Speech\" & sSpeechAPIDll

    ' 30/04/2017 Chemin plus récent
    Dim sWinSys$ = Environment.GetFolderPath(Environment.SpecialFolder.System)
    Dim sWinSysX86$ = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86)
    Dim sCheminSapi32$ = sWinSysX86 & "\Speech\Common\sapi.dll"
    Dim sCheminSapi64$ = sWinSys & "\Speech\Common\sapi.dll"
    If bFichierExiste(sCheminSapi64) Then
        sCheminSAPI = sCheminSapi64
    ElseIf bFichierExiste(sCheminSapi32) Then
        sCheminSAPI = sCheminSapi32
    End If

    ' Simple vérification de la présence de la dll dans ce cas
    ' sLHTTS3000FrClsID
    If Not bVerifierDllActiveX_InstExe(sSpeechAPI, "", _
        sSpeechAPIInst, sSpeechAPIInstTxt, sSpeechAPI_URL, sCheminSAPI) Then _
        GoTo MsgReinst

    If Not bInitMSAgent(bMSAgent, bVoixActive) Then Return False
    Return True

MsgReinst:
    MsgBox("Veuillez retester l'installation après l'installation du composant", _
        MsgBoxStyle.Information, sTitreMsg)
    Return False

End Function

Public Function bInitMSAgent( _
    ByRef bMSAgent As Boolean, ByRef bVoixActive As Boolean) As Boolean

    m_oAgent = Nothing
    m_oMerlin = Nothing

    Try
        If Not bCreerObjet(m_oAgent, sMSAgentCtrl2ProgID) Then Return False

        m_oAgent.Connected = True ' Nécessaire en DotNet !
        Const sCleAgent$ = "Merlin"
        m_oAgent.Characters.Load(sCleAgent)
        ' On peut aussi préciser le fichier .acs : .Load("Merlin", "Merlin.acs")

        m_oMerlin = m_oAgent.Characters(sCleAgent)
        'Dim iLangID% = m_oMerlin.LanguageID ' 1036 : Français
        m_oMerlin.Show()

        ' Zoom 2
        'm_oMerlin.Height *= 2
        'm_oMerlin.Width *= 2

        bMSAgent = True
        bVoixActive = True

        Return True

    Catch ex As Exception
        AfficherMsgErreur(ex, "bInitMSAgent")
        Return False

    End Try

End Function

#End Region

Public Function bMasquerMSAgent() As Boolean

    Try
        If IsNothing(m_oMerlin) Then Return False
        m_oMerlin.Hide()
        Return True
    Catch ex As Exception
        Return False
    End Try

End Function

Public Function bDireMSAgent(ByRef sParole$) As Boolean

    'If Not glb1.bMSAgent Then Exit Function
    'If Not glb1.bVoixActive Then Exit Function
    If m_oAgent Is Nothing Then Return False
    If m_oMerlin Is Nothing Then Return False

    Dim MerlinRequest As Object 'IAgentCtlRequest : Liaison anticipée
    MerlinRequest = m_oMerlin.Speak(sParole)

    ' Il faut pouvoir récupérer l'événement : compliqué
    'm_oMerlin.Wait(MerlinRequest)
    ' Autre solution : récupérer l'événement m_oAgent_RequestComplete

    ' Autre solution : + simple
    Dim iStatut% = 0 ' 2 = pending, 4 = in progress 
    Do
        iStatut = MerlinRequest.Status
        Application.DoEvents()
    Loop While iStatut = 2 Or iStatut = 4

    bDireMSAgent = True

End Function

'Private Sub m_oAgent_RequestComplete(oRequest As Object)
'
'    ' Attendre que l'agent ait finit de parler pour écrire sa réponse
'    If m_oMerlinRequest Is Nothing Then Exit Sub
'    If oRequest <> m_oSpeakRequest Then Exit Sub
'    ' Ecrire la réponse maintenant...
'
'End Sub

End Module