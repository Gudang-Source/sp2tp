VERSION 5.00
Object = "{14BE5479-3D4E-41BE-AF51-F7B42E0FA052}#114.0#0"; "vbComCtl.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Begin VB.Form frmTransGK 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Transaksi Kegiatan"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton newBtn 
      Caption         =   "Data Bulan atau Tahun Lain"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton nextBtn 
      Caption         =   "Puskesmas Berikutnya >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jumlah Kunjungan"
      Height          =   1575
      Left            =   3240
      TabIndex        =   17
      Top             =   2160
      Width           =   3015
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   3
         Left            =   3720
         TabIndex        =   11
         Top             =   1920
         Width           =   945
         _Version        =   524288
         _ExtentX        =   1667
         _ExtentY        =   529
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         SpinButtons     =   0
      End
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Width           =   945
         _Version        =   524288
         _ExtentX        =   1667
         _ExtentY        =   529
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         SpinButtons     =   0
         DisableSpins    =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   945
         _Version        =   524288
         _ExtentX        =   1667
         _ExtentY        =   529
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         SpinButtons     =   0
         DisableSpins    =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   2
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   945
         _Version        =   524288
         _ExtentX        =   1667
         _ExtentY        =   529
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Enabled         =   0   'False
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         SpinButtons     =   0
         DisableSpins    =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Perempuan"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Laki-Laki"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.ComboBox cmbData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      ItemData        =   "frmTransGK.frx":0000
      Left            =   120
      List            =   "frmTransGK.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ListBox lstData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1200
      ItemData        =   "frmTransGK.frx":0004
      Left            =   120
      List            =   "frmTransGK.frx":0006
      TabIndex        =   7
      Top             =   2520
      Width           =   3015
   End
   Begin vbComCtl.ucFrame ucFrame1 
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3413
      BeginProperty Font {6A56621B-DFAD-4DCB-A591-550817A80509} 
         Source          =   0
         Name            =   "Tahoma"
         Object.Height          =   -11
         Weight          =   700
         Underline       =   1
         Charset         =   1
         PitchFam        =   16
      EndProperty
      Caption         =   "Entry Data Gizi dan KIA"
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1800
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   4560
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbData 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmTransGK.frx":0008
         Left            =   1800
         List            =   "frmTransGK.frx":0030
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdBtn 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Puskesmas"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah T.T"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Jumlah Pelapor"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   3600
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmTransGK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim rsFind As New ADODB.Recordset
Dim rsFind2 As New ADODB.Recordset
Dim strCon As String, tEdit As Boolean
Dim kdKegiatan() As String
Dim jnsCount As Integer, kegCount As Integer
Public Ask As Variant
Public noTrans As String
Public swShowTr As Boolean

Option Explicit

Private Sub cmbData_Click(Index As Integer)
 Dim i As Integer, kdHeader As String
 
 Select Case Index
 Case 1 'Pilih Jenis Kegiatan
  lstData.Clear
  If Trim(cmbData(0).Text) <> vbNullString And _
     Trim(txtData(0).Text) <> vbNullString Then
   'Select Case cmbData(1).ListIndex
   'Case 0
    Set rsFind = con.Execute("select kode from tbHeaderKegiatan " & _
                "where nama='" & Trim(cmbData(1).Text) & "'")
    If Not rsFind.EOF Then
     kdHeader = rsFind(0).Value
    Else
     kdHeader = vbNullString
    End If
    Set rsFind = Nothing
    
    Set rsFind = con.Execute("select kode,nama from tbGKIA " & _
                 "where kdHeader='" & kdHeader & "'")
    If Not rsFind.EOF Then
     i = 0
     rsFind.MoveFirst
     While Not rsFind.EOF
      rsFind.MoveNext
      i = i + 1
     Wend
     rsFind.MoveFirst
     ReDim kdKegiatan(i) As String
     i = 0
     While Not rsFind.EOF
      lstData.AddItem rsFind(0).Value & Space(5) & rsFind(1).Value
      kdKegiatan(i) = rsFind(0).Value
      rsFind.MoveNext
      i = i + 1
     Wend
    End If
    Set rsFind = Nothing
   'End Select
  End If
 End Select
End Sub

Private Sub cmbData_GotFocus(Index As Integer)
 Select Case Index
 Case 1
  If cmbData(1).Text <> vbNullString Then
   pvnum(0).Enabled = True
   pvnum(1).Enabled = True
  End If
 End Select
End Sub

Private Sub cmbData_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmbData_LostFocus(Index As Integer)
 Dim no As Variant, MySql As String
 Dim i As Integer, nul As String
 
 Select Case Index
 Case 0
  If Trim(cmbData(0).Text) <> vbNullString And _
     Trim(txtData(0).Text) <> vbNullString And _
     noTrans = vbNullString Then
    
   noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
   noTrans = noTrans & Right(txtData(0).Text, 2) & "-"
   
   Set rsFind = con.Execute("select * from tbTransGK " & _
                "where left(no_trans,12)='" & noTrans & _
                "' order by no_trans desc")
   If Not rsFind.EOF Then
    rsFind.MoveFirst
    no = val(Right(rsFind.Fields(0).Value, 5)) + 1
    For i = 1 To 5 - Len(no)
     nul = nul & "0"
    Next
    noTrans = noTrans & nul & no
   Else
    noTrans = noTrans & "00001"
   End If
   Set rsFind = Nothing
   noTrans = Right(txtData(1).Text, 6) & "-" & noTrans
   
   If appTipe = "0" Or appTipe = "2" Then
    'txtData(1).Enabled = True
    txtData(3).Enabled = True
    txtData(4).Enabled = True
   Else
    'txtData(1).Enabled = False
    txtData(3).Enabled = False
    txtData(4).Enabled = False
   End If
   cmbData(1).Enabled = True
   lstData.Enabled = True
      
   Ask = MsgBox("Data LB3 Baru ?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "DINKES")
   If Ask = vbYes Then
    tEdit = False
    clearForm
    con.BeginTrans
    If appTipe <> "1" Then
     con.Execute "insert into tbTransGK values('" & noTrans & _
        "'," & cmbData(0).ListIndex + 1 & "," & val(txtData(0).Text) & ",'',0,'')"
    Else
     con.Execute "insert into tbTransGK values('" & noTrans & _
        "'," & cmbData(0).ListIndex + 1 & _
        "," & val(txtData(0).Text) & ",'" & kdPus & _
        "'," & txtData(3).Text & ",'" & txtData(4).Text & "')"
    End If
    con.CommitTrans
    If appTipe <> "1" Then txtData(1).SetFocus
   ElseIf Ask = vbNo Then
    tEdit = True
    noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
    noTrans = noTrans & Right(txtData(0).Text, 2) & "-"
    noTrans = Right(txtData(1).Text, 6) & "-" & noTrans
    
    MySql = "select kdPuskesmas,namaPus,jumlahTT,pelapor,no_trans from QDistincTransPusG " & _
            "where left(no_trans,5)='" & noTrans & "'"
    Set rsFind2 = con.Execute(MySql)
    If Not rsFind2.EOF Then
     If appTipe <> "1" Then
      ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              App.Path & "\dinkes07.mdb", MySql, 1, 2, 3, 4
      txtData(1).Text = Scatter_Code
      txtData(2).Text = Scatter_Code1
      txtData(3).Text = Scatter_Code2
      txtData(4).Text = Scatter_Code3
     Else
      Set rsFind = con.Execute(MySql)
      Scatter_Code4 = rsFind("no_trans")
      Set rsFind = Nothing
     End If
     noTrans = Scatter_Code4
     cmbData(1).SetFocus
     If appTipe <> "1" Then nextBtn.Enabled = True
    Else
     noTrans = vbNullString
    End If
    Set rsFind2 = Nothing
   Else     'Cancel = Close Form
    tEdit = False
    noTrans = vbNullString
    Call cmdBtn_Click(0)
   End If
  End If
  
 Case 1
   If Trim(cmbData(0).Text) <> vbNullString And _
      Trim(txtData(0).Text) <> vbNullString And _
      Ask = vbYes Then
      
    On Error Resume Next
    con.BeginTrans
    con.Execute "insert into tbTransDtlGK values('" & _
       noTrans & "','',0,0,'',0)"
    con.CommitTrans
   End If
 End Select
End Sub

Private Sub cmdBtn_Click(Index As Integer)
 Select Case Index
 Case 0
  con.Execute "delete from tbTransGK " & _
     "where kdPuskesmas=''"
  Ask = vbNo
  Call newBtn_Click
  Unload Me
 End Select
End Sub

Private Sub Form_Load()
 Dim i As Integer
 
 'strCon = BukaKoneksi
 Set con = New ADODB.Connection
 con.CursorLocation = adUseClient
 con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & App.Path & "\dinkes07.mdb;"
 
 Set rsFind = con.Execute("select nama from tbHeaderKegiatan " & _
              "where hdr='2'")
 If Not rsFind.EOF Then rsFind.MoveFirst
 While Not rsFind.EOF
  cmbData(1).AddItem rsFind(0).Value
  rsFind.MoveNext
 Wend
 Set rsFind = Nothing
 
 cmbData(0).ListIndex = 0
 If Not swShowTr Then
  noTrans = vbNullString
 End If
 swShowTr = True
 
 If appTipe = "1" Then
  txtData(1).Text = kdPus
  txtData(2).Text = nmPus
  Set rsFind = con.Execute("select juml_tt,jml_pelapor " & _
               "from tbPuskesmas where kode='" & kdPus & "'")
  If Not rsFind.EOF Then
   txtData(3).Text = rsFind.Fields(0)
   txtData(4).Text = rsFind.Fields(1)
  End If
  Set rsFind = Nothing
 End If
 
 For i = 0 To 2
  pvnum(i).ValueInteger = 0
 Next
 nextBtn.Enabled = False
End Sub

Private Sub lstData_Click()
 Dim i As Integer
 
 If Trim(cmbData(0).Text) <> vbNullString And _
    Trim(txtData(0).Text) <> vbNullString And _
    lstData.ListCount > 0 Then
 
  If Ask = vbYes Then
   con.BeginTrans
   con.Execute "update tbTransDtlGK set " & _
               "kdsubKegiatan='" & kdKegiatan(lstData.ListIndex) & _
               "' where no_Trans='" & noTrans & "'"
   con.CommitTrans
   Ask = vbNo
   If appTipe <> "1" Then nextBtn.Enabled = True
  ElseIf Ask = vbNo Then
   Set rsFind = con.Execute("select * from tbTransDtlGK " & _
                "where no_trans='" & noTrans & "' and " & _
                "kdSubKegiatan='" & kdKegiatan(lstData.ListIndex) & "'")
   If rsFind.EOF Then
    con.BeginTrans
    con.Execute "insert into tbTransDtlGK values('" & _
       noTrans & "','" & kdKegiatan(lstData.ListIndex) & "',0," & _
       "0,'',0)"
    con.CommitTrans
   End If
   Set rsFind = Nothing
  End If
 
  For i = 0 To 2
   pvnum(i).ValueInteger = 0
   pvnum(i).Enabled = True
  Next
 
    'Set rsFind = con.Execute("select usia1,usia2,usia3,usia4,usia5,usia6," & _
               "usia7,usia8,usia9,usia10,usia11,usia12 from tbPenyakit " & _
               "where kode='" & kdPenyakit(lstData.ListIndex) & "'")
    'If Not rsFind.EOF Then
    'For i = 0 To 11
    ' If Not rsFind(i).Value Then
    '     pvnum(i * 2).Enabled = False
    '     pvnum((i * 2) + 1).Enabled = False
    '    End If
    '   Next
    '  End If
    '  Set rsFind = Nothing
  
  Set rsFind = con.Execute("select jumlahL,jumlahP,total from qTransGKIA " & _
               "where kdSubKegiatan='" & kdKegiatan(lstData.ListIndex) & _
               "' and no_trans='" & noTrans & "'")
  If Not rsFind.EOF Then
   For i = 0 To 2
    pvnum(i).ValueInteger = rsFind(i).Value
   Next
  End If
  Set rsFind = Nothing
 End If
End Sub

Private Sub lstData_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub newBtn_Click()
 Dim i As Byte
 
 clearForm
 For i = 2 To 4
  txtData(i).Enabled = False
 Next
 cmbData(1).Enabled = False
 lstData.Enabled = False
 
 txtData(0).Text = vbNullString
 cmbData(0).Enabled = True
 txtData(0).Enabled = True
 cmbData(0).SetFocus
 'Ask = ""
 noTrans = vbNullString
 newBtn.Enabled = False
 nextBtn.Enabled = False
End Sub

Private Sub nextBtn_Click()
 Dim no As Variant
 Dim i As Integer, nul As String
 
 If Not tEdit Then
  noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
  noTrans = noTrans & Right(txtData(0).Text, 2) & "-"
   
  Set rsFind = con.Execute("select * from tbTransGK " & _
              "where left(no_trans,5)='" & noTrans & _
              "' order by no_trans desc")
  If Not rsFind.EOF Then
    rsFind.MoveFirst
    no = val(Right(rsFind.Fields(0).Value, 5)) + 1
    For i = 1 To 5 - Len(no)
     nul = nul & "0"
    Next
    noTrans = noTrans & nul & no
  Else
    noTrans = noTrans & "00001"
  End If
  Set rsFind = Nothing

  'Ask = vbYes
  clearForm
  con.BeginTrans
  con.Execute "insert into tbTransGK values('" & noTrans & _
     "'," & cmbData(0).ListIndex + 1 & "," & val(txtData(0).Text) & ",'',0,'')"
  con.CommitTrans
 End If
 'nextBtn.Enabled = False
 txtData(1).Text = vbNullString
 txtData(2).Text = vbNullString
 txtData(1).SetFocus
End Sub

Private Sub pvnum_Change(Index As Integer)
 Select Case Index
 Case 0, 1
  pvnum(2).Text = pvnum(0).ValueInteger + pvnum(1).ValueInteger
 End Select
End Sub

Private Sub pvnum_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub pvnum_lostFocus(Index As Integer)
 Select Case Index
 Case 0 To 2
  con.Execute "update tbTransDtlGK set " & _
              "jumlahL=" & pvnum(0).ValueInteger & ",jumlahP=" & pvnum(1).ValueInteger & _
              ",total=" & pvnum(2).ValueInteger & _
              " where no_trans='" & noTrans & "' AND " & _
              "kdSubKegiatan='" & kdKegiatan(lstData.ListIndex) & "'"
 End Select
 If Index = 2 Then lstData.SetFocus
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim MySql As String
  
  If KeyAscii = 13 Then
    Select Case Index
    Case 0
     'Call cmbData_LostFocus(0)
'     If noTrans <> vbNullString Then
'      cmbData(0).Enabled = False
'      txtData(0).Enabled = False
'      newBtn.Enabled = True
'     Else
      'txtData(1).Enabled = False
'      txtData(3).Enabled = False
'      txtData(4).Enabled = False
'      cmbData(1).Enabled = False
'      lstData.Enabled = False
'     End If
    Case 1
     If Trim(txtData(1).Text) = vbNullString Then
      If Not tEdit Then
       If appTipe <> "1" Then
        MySql = "select kode,nama from tbPuskesmas"
        ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
              App.Path & "\dinkes07.mdb", MySql, 1
        txtData(1).Text = Scatter_Code
        txtData(2).Text = Scatter_Code1
       End If
       SendKeys vbTab
       
       Call cmbData_LostFocus(0)
       If noTrans <> vbNullString Then
        cmbData(0).Enabled = False
        txtData(0).Enabled = False
        newBtn.Enabled = True
       Else
        'txtData(1).Enabled = False
        txtData(3).Enabled = False
        txtData(4).Enabled = False
        cmbData(1).Enabled = False
        lstData.Enabled = False
       End If
       
      ElseIf tEdit Then
       noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
       noTrans = noTrans & Right(txtData(0).Text, 2) & "-"
       noTrans = Right(txtData(1).Text, 6) & "-" & noTrans
   
       MySql = "select kdPuskesmas,namaPus,jumlahTT,pelapor,no_trans from qDistincTransPusG " & _
               "where left(no_trans,5)='" & noTrans & "'"
       Set rsFind2 = con.Execute(MySql)
       If Not rsFind2.EOF Then
        If appTipe <> "1" Then
          ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
             App.Path & "\dinkes07.mdb", MySql, 1, 2, 3, 4
          txtData(1).Text = Scatter_Code
          txtData(2).Text = Scatter_Code1
          txtData(3).Text = Scatter_Code2
          txtData(4).Text = Scatter_Code3
        Else
         Set rsFind = con.Execute(MySql)
         Scatter_Code4 = rsFind("no_trans")
         Set rsFind = Nothing
        End If
        noTrans = Scatter_Code4
       Else
        noTrans = vbNullString
       End If
       Set rsFind2 = Nothing
      End If
     End If
    Case Else
     SendKeys vbTab
    End Select
  End If
End Sub

Private Sub txtData_LostFocus(Index As Integer)
 Select Case Index
 Case 1
  If Ask = vbYes Then
   con.Execute "update tbTransGK set " & _
               "kdPuskesmas='" & txtData(1).Text & _
               "' where no_Trans='" & noTrans & "'"
  End If
 Case 3
  con.Execute "update tbTransGK set " & _
              "jumlahTT=" & val(txtData(3).Text) & _
              " where no_Trans='" & noTrans & "'"
 Case 4
  con.Execute "update tbTransGK set " & _
              "pelapor='" & txtData(4).Text & _
              "' where no_Trans='" & noTrans & "'"
 End Select
End Sub

Private Sub clearForm()
 Dim i As Integer
 
 For i = 0 To 2
  pvnum(i).ValueInteger = 0
  pvnum(i).Enabled = False
 Next
 
 If appTipe <> "1" Then
  For i = 3 To 4
   txtData(i).Text = vbNullString
  Next
 End If
End Sub
