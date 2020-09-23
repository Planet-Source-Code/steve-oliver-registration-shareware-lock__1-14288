VERSION 5.00
Begin VB.Form frmRegKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Program  Locked"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Unlock"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Activation Key"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Registration Code"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmRegKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a simple example of a Registration Lock
'It uses your hard drive serial number to make a key
'This way no 2 keys are alike
'The encryption used was found on www.planet-source-code.com
'-Steve Oliver (aka. genocide) =)

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Dim volbuf$, sysname$, serialnum&, sysflags&, componentlength&, res&
Dim EncryptKey As String

Private Sub Command1_Click()
'This is where they press the Unlock button, if the code is correct, it loads
'the next form in your project, if not it closes.
If TDecrypt(Text2) <> serialnum& & EncryptKey Then
    MsgBox "Invalid Activation Key", vbSystemModal
    Unload Me
    End
Else
'If code is correct, this saves an encrypted number to there registry so you only need to
'enter an unlock code once
    Call SaveSetting(App.EXEName, "Auth", "Key", TEncrypt(serialnum& & EncryptKey))

'Load the next Form of your project here
    MsgBox "Correct Activation Key", vbSystemModal
'frmMain.Show Load your main form
    Unload Me
End If
End Sub

Private Sub Form_Load()
'Encryption Key, add whatever you want for a personalized encryption
EncryptKey = "123456789"
On Error Resume Next
'This gets the unlock code from there registry if they already unlocked it
checkregd = GetSetting(App.EXEName, "Auth", "Key", "")
If checkregd <> "" Then
    checkregd = TDecrypt(checkregd)
End If
'This is an API call to get there hard drive serial number
volbuf$ = String$(256, 0)
sysname$ = String$(256, 0)
res = GetVolumeInformation("C:\", volbuf$, 255, serialnum&, _
        componentlength, sysflags, sysname$, 255)

'this is what they get for an unlock code
serialnumb = TEncrypt(serialnum& & EncryptKey)
Text1 = serialnumb

'This compares the registry to the current serial number, if they already unlocked it
'they never even see this unlock form

If checkregd = serialnum& & EncryptKey Then
'Put what you want to load after this here
    'Removes the Key, this is just for this example
    'remove this part in your actual program
    mess = MsgBox("You have already entered the Correct Unlock Code" & vbCrLf & vbCrLf & "Would you like to Remove the key and Lock it again?", vbSystemModal + vbYesNo)
    If mess = 6 Then
          Call DeleteSetting(App.EXEName, "Auth", "Key")
    End If
'frmMain.Show Load your main form here
Unload Me
End If
End Sub


Private Sub Text1_Click()
'This just highlights the text when they click the box and puts it in the clipboard
Clipboard.Clear
Clipboard.SetText Text1
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub
