VERSION 5.00
Begin VB.Form Frm_admin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Administrator"
   ClientHeight    =   3615
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5055
   ControlBox      =   0   'False
   FillColor       =   &H00800000&
   Icon            =   "Frm_admin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2135.862
   ScaleMode       =   0  'User
   ScaleWidth      =   4746.373
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      MouseIcon       =   "Frm_admin.frx":24A2
      MousePointer    =   99  'Custom
      Picture         =   "Frm_admin.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to submit"
      Top             =   2520
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      MouseIcon       =   "Frm_admin.frx":2B66
      MousePointer    =   99  'Custom
      Picture         =   "Frm_admin.frx":2CB8
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel to Abort"
      Top             =   2520
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   4815
      Begin VB.TextBox txt_pass2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "Administrator password"
         Top             =   720
         Width           =   2685
      End
      Begin VB.TextBox txt_pass1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         ToolTipText     =   "Administrator password"
         Top             =   240
         Width           =   2685
      End
      Begin VB.Label lbl_pass2 
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label lbl_pass1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   240
      Picture         =   "Frm_admin.frx":3238
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Catatan : Admin dapat melakukan konfigurasi melalui  form 'Admin setting'."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Batal"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   3210
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   3210
      Width           =   975
   End
   Begin VB.Label lbl_admin 
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan password untuk Administrator, password ini akan digunakan untuk login sebagai Administrator."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Frm_admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String
Dim Adconnnection As ADODB.Connection
Private Sub cmdCancel_Click()
Me.Hide
Unload Me
End Sub
Private Sub cmdOK_Click()
If txt_pass1.Text = "" Then
MsgBox "Please enter password.", vbInformation, "Password missing"
ElseIf txt_pass2.Text = "" Then
MsgBox "Please enter password confirm.", vbInformation, "Password missing"
ElseIf txt_pass1.Text <> txt_pass2.Text Then
MsgBox "May be typing mistake please verify the password.", vbInformation, "Password missing"
txt_pass2.Text = ""
txt_pass1.Text = ""
txt_pass1.SetFocus
Else
    If MsgBox("This password will be use as Administrator level security,Are you sure you want keep this password ?", vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
     str = "UPDATE Custom SET Pass='" & Trim(txt_pass1.Text) & "'"
    Adconnnection.Execute str
    MsgBox "Adminster can configure library settings from menu Administrator/settings.", vbInformation, "Administrator settings"
    Call globalload
    DoEvents
    Me.Hide
    Unload Me
    Exit Sub
    Else
    txt_pass2.Text = ""
    txt_pass1.Text = ""
    txt_pass1.SetFocus
    Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo errlable
  Set Adconnnection = New ADODB.Connection
  Adconnnection.CursorLocation = adUseClient
  Adconnnection.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & App.Path & "\Database\Library.mdb;Jet OLEDB:Database Password=Library;"
Exit Sub
errlable:
MsgBox Err.Description & Err.Number
End Sub

