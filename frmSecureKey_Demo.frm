VERSION 5.00
Begin VB.Form frmSecureKey_Demo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SecureKey - Demo"
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecureKey_Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraKeyGen 
      Caption         =   "KeyGen Example"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   3225
      Begin VB.TextBox txtKey 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Text            =   "Click 'Generate' to create a key."
         Top             =   255
         Width           =   3000
      End
      Begin VB.CommandButton cmdValidate 
         Caption         =   "Validate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1005
         TabIndex        =   2
         Top             =   615
         Width           =   1005
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2100
         TabIndex        =   1
         Top             =   615
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmSecureKey_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeyGen As New KeyGen 'Create an instance of the class
'Class and Demo created by Pio; ydj@aol.com; http://p-soft.shockrock.net



Private Sub cmdGenerate_Click()
    txtKey.Text = KeyGen.MakeKey 'Display the made key in txtKey
End Sub

Private Sub cmdValidate_Click()
    If KeyGen.ValidKey(txtKey.Text) Then 'Check Validation of the key
        MsgBox "Key is valid.", vbInformation + vbOKOnly, "SecureKey"
    Else
        MsgBox "Key is invalid.", vbCritical + vbOKOnly, "SecureKey"
    End If
End Sub
