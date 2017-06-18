
' Fichier modDepart.vb : Module de départ
' --------------------

Module modDepart

Public Const sNomAppli$ = "IAVB3"
Public Const sTitreMsg$ = sNomAppli
Public Const sDateVersionAppli$ = "18/06/2017"

Public ReadOnly sVersionAppli$ = _
    My.Application.Info.Version.Major & "." & _
    My.Application.Info.Version.Minor & My.Application.Info.Version.Build

#If DEBUG Then
    Public Const bDebug As Boolean = True
    'Public Const bRelease As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    'Public Const bRelease As Boolean = True
#End If

Public Sub Main()

    If bDebug Then Depart() : Exit Sub

    Try
        Depart()
    Catch ex As Exception
        AfficherMsgErreur(ex, "Main " & sTitreMsg)
    End Try

End Sub

Private Sub Depart()

    'Dim bVoixActive As Boolean
    'Dim bMSAgent As Boolean
    'If Not Is64BitProcess() Then VerifierInstallAgent(bMSAgent, bVoixActive)

    Try
        Dim oFrm As New IAVB.frmIAVB
        'oFrm.m_bMSAgent = bMSAgent
        'oFrm.m_bVoixActive = bVoixActive
        ' ShowDialog ne fonctionne pas si la session n'est pas ouverte
        'oFrm.ShowDialog()
        Application.Run(oFrm)
    Catch Ex As Exception
        AfficherMsgErreur(Ex, "Depart " & sTitreMsg)
    End Try

End Sub

End Module