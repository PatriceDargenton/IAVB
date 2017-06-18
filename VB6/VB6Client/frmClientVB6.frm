VERSION 5.00
Object = "{93774BE0-6DB5-455D-B482-F52EEEF679A7}#6.0#0"; "IAVB.ocx"
Begin VB.Form frmClientVB6 
   Caption         =   "frmTestComDotNet"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin IAVBOCX.ucIAVB ucIAVB1 
      Height          =   3495
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6165
   End
   Begin VB.TextBox tbQuestion 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdTestClientVB6 
      Caption         =   "Test client VB6"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblReponse 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   5175
   End
End
Attribute VB_Name = "frmClientVB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTestClientVB6_Click()

    Me.lblReponse.Caption = Me.ucIAVB1.strCommuniquer(Me.tbQuestion.Text)

End Sub
