Attribute VB_Name = "modUtil"

' Fichier modUtil.bas : Module de fonctions utilitaires
' -------------------

Option Explicit

Public Sub CopierPressePapier(ByVal sInfo$)

    ' Copier des informations dans le presse-papier de Windows

    If bTrapErr Then On Error GoTo Erreur
    Clipboard.Clear
    Const CF_TEXT& = 1
    Clipboard.SetText sInfo, CF_TEXT
    Exit Sub
        
Erreur:
    ' Le presse-papier peut �tre indisponible
    AfficherMsgErreur Err, "CopierPressePapier"

End Sub

Public Sub Sablier(Optional bDesactiver As Boolean = False)

    If bDesactiver Then
        Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbHourglass
    End If

End Sub

Public Function bVerifierInstallObjet(sClasse$) As Boolean

    ' V�rifier si un composant ActiveX est install�
    '  (s'il peut �tre instanci� sans erreur)
    
    Sablier
    
    Dim oObjetQcq As Object
    'Dim bObjDejaOuvert As Boolean
    ' V�rifier d'abord s'il est d�j� ouvert, pour gagner du temps
    On Error Resume Next
    Set oObjetQcq = GetObject(, sClasse)
    If Err = 0 Then
        ' Pas d'erreur, l'objet est d�j� ouvert, il est donc install�
        'bObjDejaOuvert = True
        bVerifierInstallObjet = True
        ' Quitter en laissant l'objet tel quel
        ' (fonctionne en VBA/VB6 mais pas en dotnet)
        GoTo Fin2
    End If
    
    ' S'il n'est pas d�j� ouvert, essayer de l'instancier
    'bObjDejaOuvert = False
    
    Dim bObjOuvert As Boolean
    
    Err.Clear
    Set oObjetQcq = CreateObject(sClasse)
    
    ' Il n'est pas install�
    If Err <> 0 Then Err.Clear: GoTo Fin
    
    ' Il est install�
    bObjOuvert = True
    bVerifierInstallObjet = True

Fin:
    If bObjOuvert Then
        ' Cela ne suffit pas de faire Nothing, il faut quitter avant
        On Error Resume Next
        oObjetQcq.Quit
    End If
    Set oObjetQcq = Nothing
    Err.Clear
    On Error GoTo 0
    
Fin2:
    Sablier bDesactiver:=True

End Function

Public Function bCreerObjet(ByRef oObjetQcq As Object, sClasse$) As Boolean

    ' Attention, avec Outlook, le CreateObject fait plut�t un GetObject
    '  (si l'appli �tait d�j� ouverte, elle disparait), voir bCreerObjet2

    On Error Resume Next
    Set oObjetQcq = CreateObject(sClasse)
    If Err <> 0 Then
        AfficherMsgErreur Err, "bCreerObjet", _
            "L'objet de classe [" & sClasse & "] ne peut pas �tre cr��"
        Err.Clear: Set oObjetQcq = Nothing: GoTo Fin
    End If
    bCreerObjet = True
    
Fin:
    On Error GoTo 0
    
End Function

Public Function bCreerObjet2(ByRef oObjetQcq As Object, ByVal sClasse$, _
    Optional ByRef bObjDejaOuvert As Boolean = False) As Boolean

    On Error Resume Next
    Set oObjetQcq = GetObject(, sClasse)
    If Err <> 0 Then
        bObjDejaOuvert = False
        Err.Clear
        Set oObjetQcq = CreateObject(sClasse)
    Else
        bObjDejaOuvert = True
    End If
    If Err <> 0 Then
        AfficherMsgErreur Err, "bCreerObjet2", _
            "L'objet de classe [" & sClasse & "] ne peut pas �tre cr��", _
            "Cause possible : " & sClasse & " n'est pas install�"
        Err.Clear: Set oObjetQcq = Nothing: GoTo Fin
    End If
    bCreerObjet2 = True
    
Fin:
    On Error GoTo 0
    
End Function

Public Sub AfficherMsgErreur(Erreur As Object, Optional sTitreFct$ = "", _
    Optional sInfo$ = "", Optional sDetailMsgErr$ = "")
    
    Const vbDefault% = 0
    If Screen.MousePointer <> vbDefault Then Screen.MousePointer = vbDefault
    Dim sMsg$
    If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
    If sInfo <> "" Then sMsg = sMsg & vbCrLf & sInfo
    If Erreur.Number Then
        sMsg = sMsg & vbCrLf & "Err n�" & Str$(Erreur.Number) & " :"
        sMsg = sMsg & vbCrLf & Erreur.Description
    End If
    If sDetailMsgErr <> "" Then sMsg = sMsg & vbCrLf & sDetailMsgErr
    MsgBox sMsg, vbCritical, sTitreMsg
    
End Sub



