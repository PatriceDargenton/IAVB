VERSION 5.00
Begin VB.UserControl ucIAVB 
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   ScaleHeight     =   3660
   ScaleWidth      =   5190
   Begin VB.TextBox tbEntree 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Transmettre"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblReponse 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4935
   End
End
Attribute VB_Name = "ucIAVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim iavb As New clsIAVB

Private Sub UserControl_Initialize()
    iavb.Initialiser
End Sub

Public Function strCommuniquer$(strEntree$)
    strCommuniquer = iavb.strCommuniquer(strEntree)
End Function

Private Sub cmdTest_Click()
    lblReponse.Caption = strCommuniquer(tbEntree.Text)
    'MsgBox "Disc.:" & iavb.strDiscussion
    ' La réponse vocale contient à la fois la réponse vocale
    '  et la réponse écrite (à afficher par exemple une fois la réponse vocale dite)
    'MsgBox "Réponse vocale : " & iavb.bVoixActive & " : " & iavb.strReponseVocale
End Sub
