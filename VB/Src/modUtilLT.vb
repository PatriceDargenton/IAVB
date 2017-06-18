
' Fichier modUtilLT.vb : Module de fonctions utilitaires en liaison tardive
' --------------------

Option Strict Off ' Pour oObjetQcq.Version

Module modUtilLT

' Attribut pour �viter que l'IDE s'interrompt en cas d'exception
<System.Diagnostics.DebuggerStepThrough()> _
Public Function bVerifierInstallObjet(sClasse$, _
    Optional ByRef sVersion$ = "", _
    Optional bClassID As Boolean = False, _
    Optional bLireVersion As Boolean = False) As Boolean
    'Optional ByRef sMajorVersion$ = "", _
    'Optional ByRef sMinorVersion$ = ""

    ' V�rifier si le composant est bien install�
    ' Pour les serveurs com/ActiveX mono-instance comme Outlook
    '  il faut utiliser une autre version qui teste GetObject avant 
    '  CreateObject (sinon cette fonction risque de provoquer 
    '  la fermeture du composant s'il est d�j� ouvert)

    If bClassID Then
        ' Si c'est une ClassID au lieu d'un ProgID
        '  on lit simplement la cl�
        If bCleRegistreCRExiste("TypeLib", sClasse) Then Return True
        Return False
    End If

    Dim oObjetQcq As Object = Nothing
    Try
        oObjetQcq = CreateObject(sClasse) ' sClasse = ProgID
        bVerifierInstallObjet = True
    Catch 'ex As Exception
        bVerifierInstallObjet = False
    End Try
    If bVerifierInstallObjet And bLireVersion Then
        Try
            sVersion = oObjetQcq.Version.ToString
            'sVersion = CStr(oObjetQcq.Version)
            'sMajorVersion = oObjetQcq.MajorVersion
            'sMinorVersion = oObjetQcq.MinorVersion
        Catch
        End Try
    End If
    oObjetQcq = Nothing

End Function

Public Function bCreerObjet(ByRef oObjetQcq As Object, ByRef sClasse$) As Boolean

    ' Attention, avec Outlook, le CreateObject fait plut�t un GetObject
    '  (si l'appli �tait d�j� ouverte, elle disparait), voir GetObject(, sClasse)

    Try
        oObjetQcq = CreateObject(sClasse)
        Return True
    Catch ex As Exception
        AfficherMsgErreur(ex, "bCreerObjet", "Classe de l'objet : " & sClasse)
        oObjetQcq = Nothing
        Return False
    End Try

End Function

End Module
