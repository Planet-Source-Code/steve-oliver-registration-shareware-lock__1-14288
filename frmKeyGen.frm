VERSION 5.00
Begin VB.Form frmKeyGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Key Generator"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get Key"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmKeyGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
firs = TDecrypt(Text1)
serialnumb = TEncrypt(firs)
Text2 = serialnumb
Clipboard.Clear
Clipboard.SetText Text2
End Sub

Private Sub Text2_Click()
Text2.SelStart = 0
Text2.SelLength = Len(Text2)
End Sub
