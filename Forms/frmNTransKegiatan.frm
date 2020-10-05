VERSION 5.00
Object = "{14BE5479-3D4E-41BE-AF51-F7B42E0FA052}#114.0#0"; "vbComCtl.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Begin VB.Form frmNTransKegiatan 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Transaksi Kegiatan"
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton newBtn 
      Caption         =   "Data Bulan atau Tahun Lain"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   7920
      Width           =   4095
   End
   Begin VB.CommandButton nextBtn 
      Caption         =   "Puskesmas Berikutnya >>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   13
      Top             =   7920
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jumlah Kunjungan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   19
      Top             =   5880
      Width           =   10215
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   3
         Left            =   3720
         TabIndex        =   11
         Top             =   4080
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
         Height          =   480
         Index           =   1
         Left            =   4560
         TabIndex        =   9
         Top             =   840
         Width           =   1425
         _Version        =   524288
         _ExtentX        =   2514
         _ExtentY        =   847
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Enabled         =   0   'False
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         SpinButtons     =   0
         DisableSpins    =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   480
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   1425
         _Version        =   524288
         _ExtentX        =   2514
         _ExtentY        =   847
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   1
         Enabled         =   0   'False
         Appearance      =   0
         BorderStyle     =   1
         Alignment       =   2
         SpinButtons     =   0
         DisableSpins    =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   480
         Index           =   2
         Left            =   7680
         TabIndex        =   10
         Top             =   840
         Width           =   1425
         _Version        =   524288
         _ExtentX        =   2514
         _ExtentY        =   847
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   315
         Left            =   7680
         TabIndex        =   22
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Perempuan"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   315
         Left            =   4560
         TabIndex        =   21
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Laki-Laki"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   315
         Left            =   1560
         TabIndex        =   20
         Top             =   480
         Width           =   1290
      End
   End
   Begin VB.ComboBox cmbData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      ItemData        =   "frmNTransKegiatan.frx":0000
      Left            =   120
      List            =   "frmNTransKegiatan.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   10215
   End
   Begin VB.ListBox lstData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      ItemData        =   "frmNTransKegiatan.frx":0004
      Left            =   120
      List            =   "frmNTransKegiatan.frx":0006
      TabIndex        =   7
      Top             =   3840
      Width           =   10215
   End
   Begin vbComCtl.ucFrame ucFrame1 
      Height          =   3015
      Left            =   120
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5318
      BeginProperty Font {6A56621B-DFAD-4DCB-A591-550817A80509} 
         Source          =   0
         Name            =   "Tahoma"
         Object.Height          =   -11
         Weight          =   700
         Underline       =   1
         Charset         =   1
         PitchFam        =   16
      EndProperty
      Caption         =   "Entry Data Kegiatan"
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   5760
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   4
         Left            =   1920
         TabIndex        =   5
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   1920
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   4560
         TabIndex        =   3
         Top             =   960
         Width           =   5415
      End
      Begin VB.ComboBox cmbData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         ItemData        =   "frmNTransKegiatan.frx":0008
         Left            =   1920
         List            =   "frmNTransKegiatan.frx":0030
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2535
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
         Left            =   9960
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Puskesmas"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah T.T"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
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
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Jumlah Pelapor"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
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
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   4800
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmNTransKegiatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rsFind As New ADODB.Recordset
Dim MySql As String, kdKegiatan() As String
Public Ask As Variant
Public noTrans As String

Option Explicit

Private Sub cmbData_Click(Index As Integer)
 Dim i As Integer, kdHeader As String
 
 Select Case Index
 Case 1 'Pilih Jenis Kegiatan
  For i = 0 To 2
   pvNum(i).Enabled = False
   pvNum(i).ValueInteger = 0
  Next
  lstData.Clear
  If Trim(cmbData(0).Text) <> vbNullString And _
     Trim(txtdata(0).Text) <> vbNullString Then
    Set rsFind = con.Execute("select kode from tbHeaderKegiatan " & _
                "where nama='" & Trim(cmbData(1).Text) & "'")
    If Not rsFind.EOF Then
     kdHeader = rsFind(0).Value
    Else
     kdHeader = vbNullString
    End If
    Set rsFind = Nothing
    
    Set rsFind = con.Execute("select kode,nama from tbSubKegiatan " & _
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
      lstData.AddItem rsFind(1).Value '& Space(5) & rsFind(1).Value
      kdKegiatan(i) = rsFind(0).Value
      rsFind.MoveNext
      i = i + 1
     Wend
    End If
    Set rsFind = Nothing
  End If
 End Select
End Sub

Private Sub cmbData_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys vbTab
 End If
End Sub

Private Sub cmdBtn_Click(Index As Integer)
 Select Case Index
 Case 0
  con.Execute "delete from tbTransKegiatan " & _
     "where kdPuskesmas=''"
  con.Execute "delete from tbTransDtlKegiatan " & _
     "where kdSubKegiatan=''"
  'Ask = vbNo
  'Call newBtn_Click
  Unload Me
 End Select
End Sub

Private Sub Form_Activate()
 Me.Top = 50
 Me.Left = 50
End Sub

Private Sub Form_Load()
 Set con = New ADODB.Connection
 con.CursorLocation = adUseClient
 'con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
 '            "Data Source=" & App.Path & "\dinkes07.mdb;"
 con.Open "DSN=dinkesLab"
 cmbData(0).ListIndex = 0
 Set rsFind = con.Execute("select nama from tbHeaderKegiatan " & _
              "where hdr='4'")
 If Not rsFind.EOF Then rsFind.MoveFirst
 While Not rsFind.EOF
  cmbData(1).AddItem rsFind(0).Value
  rsFind.MoveNext
 Wend
 Set rsFind = Nothing
End Sub

Private Sub lstData_Click()
 If lstData.ListCount > 0 Then
  pvNum(0).Enabled = True
  pvNum(1).Enabled = True
  pvNum(2).Enabled = True
  pvNum(0).ValueInteger = 0
  pvNum(1).ValueInteger = 0
  pvNum(2).ValueInteger = 0
  
  Set rsFind = con.Execute("select * from tbTransDtlKegiatan " & _
               "where no_trans='" & noTrans & "' and " & _
               "kdSubKegiatan='" & kdKegiatan(lstData.ListIndex) & "'")
  If rsFind.EOF Then
   con.Execute "insert into tbTransDtlKegiatan values('" & _
       noTrans & "','" & kdKegiatan(lstData.ListIndex) & "',0," & _
       "0,'',0)"
  Else
   pvNum(0).ValueInteger = rsFind("jumlahL").Value
   pvNum(1).ValueInteger = rsFind("jumlahP").Value
   pvNum(2).ValueInteger = rsFind("total").Value
  End If
  Set rsFind = Nothing
 End If
End Sub

Private Sub lstData_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  SendKeys vbTab
 End If
End Sub

Private Sub newBtn_Click()
 Dim i As Integer
 
 nextBtn.Enabled = False
 newBtn.Enabled = False
 For i = 0 To 2
  pvNum(i).Enabled = False
  pvNum(i).ValueInteger = 0
 Next
 lstData.Clear
 lstData.Refresh
 cmbData(1).Enabled = False
 cmbData(0).Locked = False
 For i = 0 To 1
  txtdata(i).Locked = False
  txtdata(i).Text = vbNullString
 Next
 For i = 2 To 4
  txtdata(i).Text = vbNullString
  txtdata(i).Enabled = False
 Next
 cmbData(0).SetFocus
End Sub

Private Sub nextBtn_Click()
 Dim i As Integer
 
 txtdata(1).Locked = False
 For i = 1 To 4
  txtdata(i).Text = vbNullString
  If i >= 3 Then txtdata(i).Enabled = False
 Next
 For i = 0 To 2
  pvNum(i).ValueInteger = 0
  pvNum(i).Enabled = False
 Next
 lstData.Enabled = False
 lstData.Clear
 cmbData(1).Enabled = False
 txtdata(1).SetFocus
End Sub

Private Sub pvnum_GotFocus(Index As Integer)
 If Index = 3 Then
  On Error Resume Next  'Error Handler berfungsi untuk mengabaikan fokus kontrol
  lstData.SetFocus
 Else
  SendKeys "{home}+{end}"
 End If
End Sub

Private Sub pvnum_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  Select Case Index
  Case 0    'Laki-Laki
   con.Execute "update tbTransDtlKegiatan set " & _
              "jumlahL=" & pvNum(0).ValueInteger & _
              " where no_trans='" & noTrans & "' AND" & _
              " kdSubKegiatan='" & kdKegiatan(lstData.ListIndex) & "'"
   pvNum(2).ValueInteger = pvNum(0).ValueInteger + pvNum(1).ValueInteger
  Case 1    'Perempuan
   con.Execute "update tbTransDtlKegiatan set " & _
              "jumlahP=" & pvNum(1).ValueInteger & _
              " where no_trans='" & noTrans & "' AND" & _
              " kdSubKegiatan='" & kdKegiatan(lstData.ListIndex) & "'"
   pvNum(2).ValueInteger = pvNum(0).ValueInteger + pvNum(1).ValueInteger
  Case 2    'Total
   con.Execute "update tbTransDtlKegiatan set " & _
              "total=" & pvNum(2).ValueInteger & _
              " where no_trans='" & noTrans & "' AND" & _
              " kdSubKegiatan='" & kdKegiatan(lstData.ListIndex) & "'"
  End Select
  SendKeys vbTab
 End If
End Sub

Private Sub txtData_GotFocus(Index As Integer)
 SendKeys "{home}+{end}"
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  Select Case Index
  Case 0
   If Trim(txtdata(0).Text) = vbNullString Or _
      val(Trim(txtdata(0).Text)) <= 0 Then
    MsgBox "Pastikan Kolom Tahun terisi Dengan Format Yang Benar", vbOKOnly + vbInformation, "DINAS KESEHATAN"
    cmbData(0).SetFocus 'Set Fokus ke Kolom Tahun
   End If
  Case 1
    If Trim(txtdata(0).Text) <> vbNullString And _
       val(Trim(txtdata(0).Text)) > 0 Then
     If Trim(txtdata(2).Text) = vbNullString Then
      MySql = "select kode,nama from tbPuskesmas"
      'ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
      '       App.Path & "\dinkes07.mdb", MySql, 1
      ShowFind "DSN=dinkesLab", MySql, 1
      txtdata(1).Text = Scatter_Code
      txtdata(2).Text = Scatter_Code1
     End If
    
     noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
     noTrans = noTrans & Right(txtdata(0).Text, 2) & "-"
     noTrans = Right(txtdata(1).Text, 6) & "-" & noTrans
     
     If Trim(txtdata(1).Text) <> vbNullString Then
      'Editing Mode
      cmbData(0).Locked = True
      txtdata(0).Locked = True
      txtdata(1).Locked = True
      nextBtn.Enabled = True
      newBtn.Enabled = True
    
      Seleksi_Proses      'Proses Pengisian Data Dimulai
     Else
      MsgBox "Pastikan Kolom Puskesmas terisi Dengan Benar", vbOKOnly + vbInformation, "DINAS KESEHATAN"
      txtdata(0).SetFocus 'Set Fokus ke Kolom Puskesmas
     End If
    Else
     MsgBox "Pastikan Kolom Tahun Terisi Dengan Format Yang Benar", vbOKOnly + vbInformation, "DINAS KESEHATAN"
     cmbData(0).SetFocus 'Set Fokus ke Kolom Tahun
    End If
  
  Case 3
   If Trim(txtdata(3).Text) <> vbNullString Then
    con.Execute "update tbTransKegiatan set " & _
        "jumlahTT=" & val(Trim(txtdata(3).Text)) & _
        " where no_trans='" & noTrans & "'"
   Else
    MsgBox "Pastikan Kolom Jumlah T.T sudah terisi", vbOKOnly + vbInformation, "DINAS KESEHATAN"
    txtdata(1).SetFocus 'Set Fokus ke Kolom Jumlah T.T
   End If
  
  Case 4
   If Trim(txtdata(4).Text) <> vbNullString Then
    con.Execute "update tbTransKegiatan set " & _
        "pelapor=" & val(Trim(txtdata(4).Text)) & _
        " where no_trans='" & noTrans & "'"
   Else
    MsgBox "Pastikan Kolom Jumlah Pelapor sudah terisi", vbOKOnly + vbInformation, "DINAS KESEHATAN"
    txtdata(3).SetFocus 'Set Fokus ke Kolom Jumlah T.T
   End If
  End Select
  SendKeys vbTab
 End If
End Sub

Private Sub txtData_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
 Select Case Index
 Case 1
  Select Case KeyCode
  Case "8", "48" To "57", "96" To "105"
   Set rsFind = con.Execute("select nama from tbPuskesmas " & _
              "where kode='" & Trim(txtdata(1).Text) & "'")
   If Not rsFind.EOF Then
    txtdata(2).Text = rsFind.Fields(0).Value
   Else
    txtdata(2).Text = vbNullString
   End If
   Set rsFind = Nothing
  End Select
 End Select
End Sub

Private Sub Seleksi_Proses()
 Dim no As Variant, nul As String
 Dim i As Integer
 
 Ask = MsgBox("Data LB4 Baru ?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "DINAS KESEHATAN")
 If Ask = vbYes Then        'Data Baru
   Set rsFind = con.Execute("select * from tbTransKegiatan " & _
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
   
   con.Execute "insert into tbTransKegiatan values('" & noTrans & _
       "'," & cmbData(0).ListIndex + 1 & "," & val(txtdata(0).Text) & _
       ",'" & Trim(txtdata(1).Text) & "',0,0)"
   
   txtdata(3).Enabled = True
   txtdata(4).Enabled = True
   
   cmbData(1).Enabled = True
   lstData.Enabled = True
 ElseIf Ask = vbNo Then     'Data Lama
  MySql = "select kdPuskesmas,namaPus,jumlahTT,pelapor,no_trans from qDistincTransPusK " & _
          "where right(left(no_trans,11),4)='" & Mid(noTrans, 8, 4) & _
          "' and kdpuskesmas='" & txtdata(1).Text & "'"
  Set rsFind = con.Execute(MySql)
  If Not rsFind.EOF Then
    'ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    '         App.Path & "\dinkes07.mdb", MySql, 1, 2, 3, 4
    ShowFind "DSN=dinkesLab", MySql, 1, 2, 3, 4
    txtdata(3).Text = Scatter_Code2
    txtdata(4).Text = Scatter_Code3
    noTrans = Scatter_Code4
    
    cmbData(0).Locked = True
    txtdata(0).Locked = True
    txtdata(1).Locked = True
    txtdata(3).Enabled = True
    txtdata(4).Enabled = True
    cmbData(1).Enabled = True
    lstData.Enabled = True
    For i = 0 To 2
     pvNum(i).Enabled = True
    Next
    nextBtn.Enabled = True
    newBtn.Enabled = True
  Else
    MsgBox "Data Puskesmas " & txtdata(2).Text & vbCrLf & _
           "Bulan " & cmbData(0).Text & vbCrLf & _
           "Tahun " & txtdata(0).Text & vbCrLf & _
           "Belum Ada, Silahkan Isi Data Barunya    ", vbOKOnly + vbInformation, "DINAS KESEHATAN"
    txtdata(1).Locked = False
    txtdata(0).SetFocus
  End If
  Set rsFind = Nothing
 Else                       'Kembali ke Kolom Tahun
  noTrans = vbNullString
  txtdata(0).Text = vbNullString
  txtdata(1).Text = vbNullString
  txtdata(2).Text = vbNullString
  txtdata(0).Locked = False
  cmbData(0).Locked = False
  newBtn.Enabled = False
  nextBtn.Enabled = False
  pvNum(3).SetFocus 'Set Fokus ke Combo Bulan
 End If
End Sub
