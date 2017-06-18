
Imports System.Text ' Pour StringBuilder
Imports System.Text.Encoding ' Pour GetEncoding

Module modVBUtil

' Le code page 1252 correspond à FileOpen de VB .NET, l'équivalent en VB6 de
'  Open sCheminFichier For Input As #1
Private Const iCodePageWindowsLatin1252% = 1252 ' windows-1252 = msoEncodingWestern

' L'encodage UTF-8 est le meilleur compromis encombrement/capacité 
'  il permet l'encodage par exemple du grec, sans doubler la taille du texte
'(mais le décodage est plus complexe en interne et les caractères ne s'affichent
' pas bien dans les certains logiciels utilitaires comme WinDiff, 
' ni par exemple en csv pour Excel)
' http://fr.wikipedia.org/wiki/Unicode
' 65001 = Unicode UTF-8, 65000 = Unicode UTF-7
Private Const iEncodageUnicodeUTF8% = 65001

#Region "VBUtil"

Public Function sVBLeft$(sTxt$, iLeft%)

    Dim iLong% = sTxt.Length
    If iLeft > iLong Then Return sTxt
    Dim sLeft$ = sTxt.Substring(0, iLeft)
    Return sLeft

End Function

Public Function sVBRight$(sTxt$, iRight%)

    Dim iLong% = sTxt.Length
    If iRight > iLong Then Return sTxt
    Dim sRight$ = sTxt.Substring(iLong - iRight, iRight)
    Return sRight

End Function

Public Function iVBLen%(sTxt$)

    Return sTxt.Length

End Function

Public Function sVBMid$(sTxt$, iDeb%)

    Dim sMid$ = sTxt.Substring(iDeb - 1)
    Return sMid

End Function

Public Function sVBMid$(sTxt$, iDeb%, iLong%)

    Dim sMid$
    If iLong + iDeb - 1 > sTxt.Length Then
        sMid = ""
    Else
        sMid = sTxt.Substring(iDeb - 1, iLong)
    End If
    Return sMid

End Function

Public Function sVBTrim$(sTxt$)

    Dim sTrim2$ = sTrimRecursif(sTxt)
    Return sTrim2

End Function

Public Function sTrimRecursif$(sTxt$)

    ' Fonction récursive ! Attention aux performances : mauvaise idée ! 
    ' (héritage Microsoft.VisualBasic.dll)

    Dim str2$
    Dim iLong% = sTxt.Length
    If (sTxt Is Nothing) OrElse iLong = 0 Then Return ""
    Select Case sTxt.Chars(0)
    Case " "c : Return sTrimRecursif(sTxt.Substring(1))
    'Case " "c, ChrW(12288) : Return sTrimRecursif(sTxt.Substring(1))
    End Select
    Select Case sTxt.Chars(iLong - 1)
    Case " "c : Return sTrimRecursif(sTxt.Substring(0, iLong - 1))
    'Case " "c, ChrW(12288) : Return sTrimRecursif(sTxt.Substring(0, iLong - 1))
    End Select
    str2 = sTxt
    Return str2

End Function

Public Function iVBVal%(sTxt$, Optional iValDef% = 0)

    Dim iRes%
    If Integer.TryParse(sTxt, iRes) Then Return iRes
    Return iValDef

End Function

Public Function VBChr$(iCharCode%)

    Dim cChrConvert As Char = System.Convert.ToChar(iCharCode)
    Return cChrConvert

End Function

#End Region

Public Function sEnleverAccents$(sChaine$, Optional bTexteUnicode As Boolean = False)
    'Optional bMinuscule As Boolean = True

    ' Enlever les accents (voir aussi modUtilFichier.sConvNomDos)

    If sChaine.Length = 0 Then sEnleverAccents = "" : Exit Function

    Const sEncodageIso8859_15$ = "iso-8859-15"
    Const sEncodageIso8859_8$ = "iso-8859-8"
    'Const sEncodageDest$ = "windows-1252"

    ' Frédéric François, cœur
    ' iso-8859-8   -> windows-1252 : Frederic Francois, cour ' Meilleure solution
    ' windows-1251 -> windows-1252 : Frederic Francois, c?ur ' Ancienne solution
    ' iso-8859-15  -> windows-1252 : Frédéric François, c½ur ' Utile pour détecter <>

    ' Codepage 1241 = "windows-1251" = cyrillic
    ' Tableau de caractères sur 8 bit
    'Dim aOctets As Byte() = GetEncoding(1251).GetBytes(sChaine)
    ' Chaîne de caractères sur 7 bit
    'sEnleverAccents = ASCII.GetString(aOctets) ' Ok mais reste cœur qui est converti en c?ur

    Dim iEncodageDest% = iCodePageWindowsLatin1252
    If bTexteUnicode Then iEncodageDest = iEncodageUnicodeUTF8
    Dim encodage1252 As Encoding = GetEncoding(iCodePageWindowsLatin1252)
    Dim encodage8859_8 As Encoding = GetEncoding(sEncodageIso8859_8)
    Dim encodageDest As Encoding = GetEncoding(iEncodageDest)
    Dim encodageIso8859_15 As Encoding = GetEncoding(sEncodageIso8859_15)

    Dim aOctets As Byte() = encodage8859_8.GetBytes(sChaine) ' "iso-8859-8"
    sEnleverAccents = encodageDest.GetString(aOctets)        ' 1252 ou UTF8
    'If bDebug Then Debug.WriteLine("' " & sEncodageSrc & " -> " & sEncodageDest & " : " & sEnleverAccents)

    ' Détection des caractères propres à iso-8859-15 : ¤ ¦ ¨ ´ ¸ ¼ ½ ¾ € Š š Ž ž Œ œ Ÿ
    ' http://fr.wikipedia.org/wiki/ISO_8859-15
    If String.Compare(encodageIso8859_15.GetString( _
        encodage1252.GetBytes(sChaine)), sChaine) = 0 Then GoTo Fin

    Dim i% = 0
    Dim iLen% = sChaine.Length
    Dim sChaineIso$ = encodageIso8859_15.GetString(encodageDest.GetBytes(sChaine))
    Dim ac1, ac2, ac3 As Char()
    ac1 = sChaine.ToCharArray
    ac2 = sChaineIso.ToCharArray
    ac3 = sEnleverAccents.ToCharArray
    Dim sbDest As New StringBuilder
    For i = 0 To iLen - 1
        If ac1(i) = ac2(i) Then
            sbDest.Append(ac3(i))
        Else
            Select Case ac1(i) ' ¤ ¦ ¨ ´ ¸ ¼ ½ ¾ € Š š Ž ž Œ œ Ÿ
            Case "¤"c : sbDest.Append("o")
            Case "¦"c : sbDest.Append("|")
            Case "¨"c : sbDest.Append("..")
            Case "´"c : sbDest.Append("'")
            Case "¸"c : sbDest.Append(",")
            Case "¼"c : sbDest.Append("1/4")
            Case "½"c : sbDest.Append("1/2")
            Case "¾"c : sbDest.Append("3/4")
            Case "€"c : sbDest.Append("E")
            Case "Š"c : sbDest.Append("S")
            Case "š"c : sbDest.Append("s")
            Case "Ž"c : sbDest.Append("Z")
            Case "ž"c : sbDest.Append("z")
            Case "œ"c : sbDest.Append("oe")
            Case "Œ"c : sbDest.Append("OE")
            Case "Ÿ"c : sbDest.Append("Y")
            Case Else
                'If bDebug Then Debug.WriteLine("?? : " & ac1(i) & ac2(i) & ac3(i))
                sbDest.Append(ac1(i)) ' 22/05/2010 Laisser le car. si non trouvé
            End Select
        End If
    Next i
    sEnleverAccents = sbDest.ToString

Fin:
    'If bMinuscule Then sEnleverAccents = sEnleverAccents.ToLower

End Function

End Module
