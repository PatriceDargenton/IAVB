Attribute VB_Name = "modSynthVocale"

' Fichier modSynthVocale.bas : Module de synthèse vocale via MS-Agent
' --------------------------

Option Explicit

'Private WithEvents m_oAgent As Agent ' WithEvents pour recevoir l'év. _RequestComplete
'Private m_oMerlin As IAgentCtlCharacter
'Private m_oMerlinRequest As IAgentCtlRequest
Private m_oAgent As Object
Private m_oMerlin As Object

Public Function bDireMSAgent(sParole$) As Boolean
    
    If m_oAgent Is Nothing Then Exit Function
    If m_oMerlin Is Nothing Then Exit Function
    
    Dim MerlinRequest As Object 'AgentObjects.IAgentCtlRequest
    Set MerlinRequest = m_oMerlin.Speak(sParole)

    ' Il faut pouvoir récupérer l'événement : compliqué
    'm_oMerlin.Wait(MerlinRequest)
    ' Autre solution : récupérer l'événement m_oAgent_RequestComplete

    ' Autre solution : + simple
    Dim iStatut% ' 2 = pending, 4 = in progress
    Do
        iStatut = MerlinRequest.Status
        DoEvents
    Loop While iStatut = 2 Or iStatut = 4

    bDireMSAgent = True

End Function

'Private Sub m_oAgent_RequestComplete(ByVal oRequest As Object)
'
'    ' Attendre que l'agent ait finit de parler pour écrire sa réponse
'    If m_oMerlinRequest Is Nothing Then Exit Sub
'    If oRequest <> m_oSpeakRequest Then Exit Sub
'    ' Ecrire la réponse maintenant...
'
'End Sub

Public Function bInitAgent() As Boolean
    
    If bTrapErr Then On Error GoTo Erreur
    
    Set m_oAgent = Nothing
    Set m_oMerlin = Nothing
    
    If Not bCreerObjet(m_oAgent, "Agent.Control") Then Exit Function
    'Set m_oAgent = New Agent
    
    m_oAgent.Connected = True ' Nécessaire en DotNet !
    Const sCleAgent$ = "Merlin"
    m_oAgent.Characters.Load sCleAgent
    ' On peut aussi préciser le fichier .acs : .Load("Merlin", "Merlin.acs")
    
    Set m_oMerlin = m_oAgent.Characters(sCleAgent)
    m_oMerlin.Show
    
    bInitAgent = True
    Exit Function

Erreur:
    Exit Function

End Function
