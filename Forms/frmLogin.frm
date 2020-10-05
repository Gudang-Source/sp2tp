VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3570
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2109.272
   ScaleMode       =   0  'User
   ScaleWidth      =   3732.31
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3735
      Begin VB.TextBox txtData 
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
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Password"
         Top             =   720
         Width           =   2325
      End
      Begin VB.TextBox txtData 
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
         Index           =   0
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "EmployeeID  for employee"
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         TabIndex        =   11
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama User:"
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
         TabIndex        =   10
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdButton 
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
      Index           =   1
      Left            =   2760
      MouseIcon       =   "frmLogin.frx":24A2
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":25F4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel to Abort"
      Top             =   2640
      Width           =   1020
   End
   Begin VB.CommandButton cmdButton 
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
      Index           =   0
      Left            =   1680
      MouseIcon       =   "frmLogin.frx":2B74
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":2CC6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click to submit"
      Top             =   2640
      Width           =   1020
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      ItemData        =   "frmLogin.frx":3238
      Left            =   1440
      List            =   "frmLogin.frx":3242
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Enter as"
      Top             =   960
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   112.674
      X2              =   3605.552
      Y1              =   496.299
      Y2              =   496.299
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih spesifikasi sebagai apa dan masukkan password serta nama user."
      Height          =   615
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmLogin.frx":325B
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Batal"
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sebagai"
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim passOK As Boolean

Option Explicit

Private Sub cmbData_Click()
 Select Case cmbData.ListIndex
 Case 0
  txtdata(0).Enabled = False
 Case 1
  txtdata(0).Enabled = True
 End Select
End Sub

Private Sub cmdButton_Click(Index As Integer)
 Select Case Index
 Case 0
  Select Case cmbData.ListIndex
  Case 0
   If Trim(txtdata(1).Text) = "admin" Then
    passOK = True
   Else
    passOK = False
    txtdata(1).Text = vbNullString
    txtdata(1).SetFocus
    MsgBox "Password anda salah", vbExclamation + vbOKOnly, "DINKES - Login"
   End If
  Case 1
   'Under CONSTRUCTION
  End Select
  
  If passOK Then
   Unload Me
   Load startFrm
   startFrm.Show
  End If
 Case 1
  End
 End Select
End Sub

Private Sub Form_Load()
 passOK = False
 cmbData.ListIndex = 0
End Sub
