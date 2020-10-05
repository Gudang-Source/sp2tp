VERSION 5.00
Begin VB.Form frmGlobal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Global"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Caption         =   "Tutup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   8160
         MouseIcon       =   "frmGlobal.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmGlobal.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   7080
         MouseIcon       =   "frmGlobal.frx":06CC
         MousePointer    =   99  'Custom
         Picture         =   "frmGlobal.frx":081E
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save record"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   1020
         Index           =   9
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2640
         Width           =   7335
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   8
         Left            =   1920
         TabIndex        =   22
         Top             =   2280
         Width           =   7335
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   5280
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   6360
         TabIndex        =   19
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   5280
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   6360
         TabIndex        =   17
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   5280
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   6360
         TabIndex        =   15
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   5280
         TabIndex        =   12
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   6360
         TabIndex        =   11
         Top             =   600
         Width           =   2895
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   4
         ItemData        =   "frmGlobal.frx":0DB6
         Left            =   2760
         List            =   "frmGlobal.frx":0DC0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   3
         ItemData        =   "frmGlobal.frx":0DD6
         Left            =   2760
         List            =   "frmGlobal.frx":0DE0
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   2
         ItemData        =   "frmGlobal.frx":0DF6
         Left            =   2760
         List            =   "frmGlobal.frx":0E00
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   1
         ItemData        =   "frmGlobal.frx":0E16
         Left            =   2760
         List            =   "frmGlobal.frx":0E20
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   0
         ItemData        =   "frmGlobal.frx":0E36
         Left            =   2760
         List            =   "frmGlobal.frx":0E43
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lbldata 
         Height          =   300
         Left            =   1920
         TabIndex        =   25
         Top             =   2280
         Width           =   7335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama Instansi"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9360
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6360
         TabIndex        =   14
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5280
         TabIndex        =   13
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mode Entry Puskesmas"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   2220
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mode Entry Kecamatan"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2220
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mode Entry Kabupaten"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2220
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mode Entry Propinsi"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aplikasi Digunakan untuk"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exitFlag As Boolean
Dim nmInstansi As String, nmInst As String
Dim adoCon As ADODB.Connection
Dim strSQL As String

Dim kdProp As String, nmProp As String
Dim kdKab As String, nmKab As String
Dim kdKec As String, nmKec As String
Dim kdPus As String, nmPus As String

Option Explicit

Private Sub cmbData_Click(Index As Integer)
 Dim i As Integer
 
 Select Case Index
 Case 0
  If cmbData(0).ListIndex = 0 Then
   cmbData(1).ListIndex = 1: cmbData(1).Enabled = False
   cmbData(2).ListIndex = 1: cmbData(2).Enabled = False
   cmbData(3).ListIndex = 0: cmbData(3).Enabled = False
   cmbData(4).ListIndex = 0: cmbData(4).Enabled = False
  ElseIf cmbData(0).ListIndex = 1 Then
   cmbData(1).ListIndex = 1: cmbData(1).Enabled = False
   cmbData(2).ListIndex = 1: cmbData(2).Enabled = False
   cmbData(3).ListIndex = 1: cmbData(3).Enabled = False
   cmbData(4).ListIndex = 1: cmbData(4).Enabled = False
  Else
   cmbData(1).ListIndex = 0: cmbData(1).Enabled = True
   cmbData(2).ListIndex = 0: cmbData(2).Enabled = True
   cmbData(3).ListIndex = 0: cmbData(3).Enabled = True
   cmbData(4).ListIndex = 0: cmbData(4).Enabled = True
  End If
 Case 1
  If cmbData(1).ListIndex = 0 Then
   For i = 0 To 1
    txtdata(i).Enabled = False
   Next
   kdProp = vbNullString: nmProp = vbNullString
  Else
   For i = 0 To 1
    txtdata(i).Enabled = True
   Next
  End If
 Case 2
   If cmbData(2).ListIndex = 0 Then
    For i = 2 To 3
     txtdata(i).Enabled = False
    Next
    kdKab = vbNullString: nmKab = vbNullString
   Else
    For i = 2 To 3
     txtdata(i).Enabled = True
    Next
   End If
 Case 3
   If cmbData(3).ListIndex = 0 Then
    For i = 4 To 5
     txtdata(i).Enabled = False
    Next
    kdKec = vbNullString: nmKec = vbNullString
   Else
    For i = 4 To 5
     txtdata(i).Enabled = True
    Next
   End If
 Case 4
   If cmbData(4).ListIndex = 0 Then
    For i = 6 To 7
     txtdata(i).Enabled = False
    Next
    txtdata(8).Visible = True
    nmInstansi = vbNullString
    kdProp = vbNullString: nmProp = vbNullString
    nmInst = vbNullString
   Else
    For i = 6 To 7
     txtdata(i).Enabled = True
    Next
    txtdata(8).Visible = False
    nmInstansi = "PUSKESMAS "
   End If
   lbldata.Caption = nmInstansi
 End Select
End Sub

Private Sub cmdButton_Click(Index As Integer)
 Select Case Index
 Case 0     'Simpan Setting
  Set adoCon = New ADODB.Connection
  adoCon.CursorLocation = adUseClient
  adoCon.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & App.Path & "\dinkes07.mdb;"
  If MsgBox("Apakah data sudah benar ?", vbYesNo + vbQuestion, _
     "DINKES - SP2TP") = vbYes Then
   strSQL = "delete from tbsetglobal"
   adoCon.Execute strSQL
   
   'Simpan ke tabel set global
   strSQL = "insert into tbSetGlobal values('" & _
            kdProp & "','" & nmProp & "'," & _
            cmbData(1).ListIndex & ",'" & kdKab & _
            "','" & nmKab & "'," & cmbData(2).ListIndex & _
            ",'" & kdKec & "','" & nmKec & "'," & _
            cmbData(3).ListIndex & "," & cmbData(0).ListIndex & _
            ",'" & kdPus & "','" & nmPus & "'," & _
            cmbData(4).ListIndex & ",'" & nmInst & "','" & _
            txtdata(9).Text & "')"
   adoCon.Execute strSQL
   
   If cmbData(1).ListIndex = 1 Then
    strSQL = "insert into tbProp values('" & _
             kdProp & "','" & nmProp & "')"
    adoCon.Execute strSQL
   End If
   If cmbData(2).ListIndex = 1 Then
    strSQL = "insert into tbKab values('" & _
             kdKab & "','" & nmKab & "')"
    adoCon.Execute strSQL
   End If
   If cmbData(3).ListIndex = 1 Then
    strSQL = "insert into tbKec values('" & _
             kdKec & "','" & nmKec & "')"
    adoCon.Execute strSQL
   End If
   If cmbData(4).ListIndex = 1 Then
    strSQL = "insert into tbPuskesmas values('" & _
             kdPus & "','" & nmPus & "','" & _
             kdProp & "','" & kdKab & "','" & _
             kdKec & "','" & txtdata(9).Text & _
             "','',0,0)"
    adoCon.Execute strSQL
   End If
  End If
  adoCon.Close
  exitFlag = True
  Unload Me
 Case 1     'Tutup
  exitFlag = True
  Unload Me
 End Select
End Sub

Private Sub Form_Load()
 exitFlag = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Not exitFlag Then
  Cancel = True
  MsgBox "Gunakan Tombol Tutup", vbOKOnly + vbInformation, "DINKES - SP2TP"
 Else
  Cancel = False
 End If
End Sub

Private Sub txtdata_Change(Index As Integer)
 Select Case Index
 Case 0
  kdProp = txtdata(0).Text
 Case 1
  nmProp = txtdata(1).Text
 Case 2
  kdKab = txtdata(2).Text
 Case 3
  nmKab = txtdata(3).Text
 Case 4
  kdKec = txtdata(4).Text
 Case 5
  nmKec = txtdata(5).Text
 Case 6
  kdPus = txtdata(6).Text
 Case 7
  lbldata.Caption = nmInstansi & txtdata(7).Text
  nmPus = txtdata(7).Text
  nmInst = lbldata.Caption
 End Select
End Sub
