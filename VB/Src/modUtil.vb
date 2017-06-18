
' Fichier modUtil.vb : Module de fonctions utilitaires
' ---------------

Imports System.Text ' Pour Encoding

Module modUtil

Public Function Is64BitProcess() As Boolean
    Return (IntPtr.Size = 8)
End Function

Public Function bFichierExiste(sCheminFichier$, _
    Optional bPrompt As Boolean = False) As Boolean

    ' Retourne True si un fichier correspondant est trouvé
    ' Ne fonctionne pas avec un filtre, par ex. du type C:\*.txt
    bFichierExiste = IO.File.Exists(sCheminFichier)

    If Not bFichierExiste And bPrompt Then _
        MsgBox("Impossible de trouver le fichier :" & vbLf & sCheminFichier, _
            MsgBoxStyle.Critical, sTitreMsg & " - Fichier introuvable")

End Function

Public Function sDossierParent$(sCheminDossier$)

    ' Renvoyer le chemin du dossier parent
    ' Ex.: C:\Tmp\Tmp2 -> C:\Tmp

    sDossierParent = IO.Path.GetDirectoryName(sCheminDossier)

End Function

Public Sub OuvrirAppliAssociee(sCheminFichier$, _
    Optional bMax As Boolean = False, _
    Optional bVerifierFichier As Boolean = True)
    If bVerifierFichier Then
        ' Ne pas vérifier le fichier si c'est une URL à lancer
        If Not bFichierExiste(sCheminFichier, bPrompt:=True) Then Exit Sub
    End If
    Dim p As New Process
    p.StartInfo = New ProcessStartInfo(sCheminFichier)
    If bMax Then p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized
    p.Start()
End Sub

Public Function bVerifierDllActiveX_InstExe(sTitreComposant$, _
     sClasseDllActiveX$, sExeInstall$, sTypeComposant$, _
     sURLInst$, _
    Optional sCheminFichierDoitExister$ = "", _
    Optional sDossierInst$ = "\Installation", _
    Optional bClassID As Boolean = False) As Boolean

    ' Vérifier et installer le cas échéant un composant Dll ActiveX via un exe

    Dim sTitreMsg$ = sNomAppli & " : Installation de " & sTitreComposant

    Dim bOk As Boolean = False
    If sCheminFichierDoitExister.Length > 0 Then
        If bFichierExiste(sCheminFichierDoitExister) Then bOk = True
    ElseIf bVerifierInstallObjet(sClasseDllActiveX, , bClassID) Then
        bOk = True
    End If
    If bOk Then
        MsgBox("Le composant [" & sTitreComposant & "] est bien installé.", _
            MsgBoxStyle.Exclamation, sTitreMsg)
        bVerifierDllActiveX_InstExe = True
        Exit Function
    End If

    Dim sCheminExeInstall$ = Application.StartupPath & _
        sDossierInst & "\" & sExeInstall
    Dim sMsg1$ = _
        "Le composant " & sTitreComposant & vbLf & _
        "n'est pas installé sur ce poste :"

    If Not bFichierExiste(sCheminExeInstall) Then
        If MsgBoxResult.Cancel = MsgBox( _
            sMsg1 & vbLf & _
            "Cliquez sur OK pour afficher la page de téléchargement :" & vbLf & _
            sTypeComposant, _
            MsgBoxStyle.Critical Or MsgBoxStyle.OkCancel, sTitreMsg) Then _
            Application.Exit() : Return False
        OuvrirAppliAssociee(sURLInst, bVerifierFichier:=False)
        Return False
    End If

    If MsgBoxResult.Cancel = MsgBox(sMsg1 & vbLf & _
        "Cliquez sur Ok pour installer ce composant :" & vbLf & _
        sTypeComposant, _
        vbOKCancel Or vbCritical, sTitreMsg) Then Return False

    OuvrirAppliAssociee(sCheminExeInstall)

    ' Poursuivre les autres installations
    Return True

End Function

Public Function bCleRegistreCRExiste(sCle$, _
    Optional sSousCle$ = "") As Boolean

    ' Vérifier si une clé ClassesRoot existe dans la base de registre
    Try
        ' This call goes to the Catch block if the registry key is not set.
        Dim rkCle As Microsoft.Win32.RegistryKey = _
            Microsoft.Win32.Registry.ClassesRoot
        rkCle = rkCle.OpenSubKey(sCle & "\\" & sSousCle)
        If IsNothing(rkCle) Then Return False
        ' Si la version est présent, on devrait pouvoir l'obtenir ainsi :
        'string[] aryTemp = rkCle.GetSubKeyNames();
        'string sVersion = aryTemp[0];
        'aryTemp = sVersion.Split('.');
        'iMajorVer = short.Parse(aryTemp[0] ,System.Globalization.NumberStyles.AllowHexSpecifier);
        'iMinusVer = short.Parse(aryTemp[1] ,System.Globalization.NumberStyles.AllowHexSpecifier);
        Return True
    Catch
        Return False
    End Try

End Function

Public Sub AfficherMsgErreur(ByRef Ex As Exception, _
    Optional sTitreFct$ = "", Optional sInfo$ = "", _
    Optional sDetailMsgErr$ = "", _
    Optional bCopierMsgPressePapier As Boolean = True, _
    Optional ByRef sMsgErrFinal$ = "")

    If Not Cursor.Current.Equals(Cursors.Default) Then _
        Cursor.Current = Cursors.Default
    Dim sMsg$ = ""
    If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
    If sInfo <> "" Then sMsg &= vbCrLf & sInfo
    If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
    If Ex.Message <> "" Then
        sMsg &= vbCrLf & Ex.Message.Trim
        If Not IsNothing(Ex.InnerException) Then _
            sMsg &= vbCrLf & Ex.InnerException.Message
    End If
    If bCopierMsgPressePapier Then bCopierPressePapier(sMsg)
    sMsgErrFinal = sMsg
    MsgBox(sMsg, MsgBoxStyle.Critical)

End Sub

Public Function bCopierPressePapier(sInfo$, Optional ByRef sMsgErr$ = "") As Boolean

    ' Copier des informations dans le presse-papier de Windows
    ' (elles resteront jusqu'à ce que l'application soit fermée)

    Try
        Dim dataObj As New DataObject
        dataObj.SetData(DataFormats.Text, sInfo)
        Clipboard.SetDataObject(dataObj)
        Return True
    Catch ex As Exception
        ' Le presse-papier peut être indisponible
        AfficherMsgErreur(ex, "CopierPressePapier", bCopierMsgPressePapier:=False, _
            sMsgErrFinal:=sMsgErr)
        Return False
    End Try

End Function

Public Sub Attendre(Optional iMilliSec% = 200)
    Threading.Thread.Sleep(iMilliSec)
End Sub

Public Sub TraiterMsgSysteme_DoEvents()
    Application.DoEvents()
End Sub

End Module