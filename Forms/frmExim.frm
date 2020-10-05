VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExim 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export & Import XLS Data"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "frmExim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox cmbData 
         Height          =   315
         ItemData        =   "frmExim.frx":030A
         Left            =   360
         List            =   "frmExim.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin MSComDlg.CommonDialog cmDlg 
         Left            =   1680
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Export ke Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   3255
      End
      Begin VB.ComboBox cmbBulan 
         Height          =   315
         ItemData        =   "frmExim.frx":037E
         Left            =   360
         List            =   "frmExim.frx":03A6
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtTahun 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2160
         TabIndex        =   2
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Import dari Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3720
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3720
         Y1              =   1800
         Y2              =   1800
      End
   End
End
Attribute VB_Name = "frmExim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim Rs As ADODB.Recordset
Private myList1 As ListItem
Private Status As String
Dim lstSql As String, lstPil As Byte
Dim nChecked As Integer, strSQL As String
Dim nBulan As Variant, nTahun As Variant

Option Explicit

Private Sub cmdButton_Click(Index As Integer)
 Dim nmTabel As String, nmFile As String
 Dim ExistFile As Variant, nTrans As String
 
 Select Case Index
 Case 1     'Export Data
  nBulan = cmbBulan.ListIndex + 1
  nTahun = txtTahun.Text
  nmTabel = vbNullString
  nmFile = vbNullString
    
  Select Case cmbData.ListIndex
  Case 0
   strSQL = "select * from tbTransPenyakit where " & _
            "bulan=" & nBulan & " and " & _
            "tahun=" & nTahun & ""
   Set Rs = con.Execute(strSQL)
   If Not Rs.EOF Then
     'ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
     '        App.Path & "\dinkes07.mdb", strSQL, 1
     ShowFind "DSN=dinkesLab", strSQL, 1
     nTrans = Scatter_Code
   End If
   Set Rs = Nothing
   
   nmTabel = "select * from xlTransPenyakit where " & _
             "bulan=" & nBulan & " and " & _
             "tahun=" & nTahun & " and " & _
             "no_trans='" & nTrans & "'"
   cmDlg.Filter = "(*.xls)|*.xls"
   'nmFile = cmDlg.FileName
  Case 1
   strSQL = "select * from tbTransGK where " & _
            "bulan=" & nBulan & " and " & _
            "tahun=" & nTahun & ""
   Set Rs = con.Execute(strSQL)
   If Not Rs.EOF Then
    'ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    '         App.Path & "\dinkes07.mdb", strSQL, 1
    ShowFind "DSN=dinkesLab", strSQL, 1
    nTrans = Scatter_Code
   End If
   Set Rs = Nothing
   
   nmTabel = "select * from xlTransGKIA where " & _
             "bulan=" & nBulan & " and " & _
             "tahun=" & nTahun & " and " & _
             "no_trans='" & nTrans & "'"
   cmDlg.Filter = "(*.xls)|*.xls"
   'nmFile = cmDlg.FileName
  Case 2
   strSQL = "select * from tbTransKegiatan where " & _
            "bulan=" & nBulan & " and " & _
            "tahun=" & nTahun & ""
   Set Rs = con.Execute(strSQL)
   If Not Rs.EOF Then
    'ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
    '         App.Path & "\dinkes07.mdb", strSQL, 1
    ShowFind "DSN=dinkesLab", strSQL, 1
    nTrans = Scatter_Code
   End If
   Set Rs = Nothing
  
   nmTabel = "select * from xlTransKegiatan where " & _
             "bulan=" & nBulan & " and " & _
             "tahun=" & nTahun & " and " & _
             "no_trans='" & nTrans & "'"
   cmDlg.Filter = "(*.xls)|*.xls"
   'nmFile = cmDlg.FileName
  End Select
  cmDlg.ShowSave
  nmFile = cmDlg.FileTitle
  If nmTabel <> vbNullString And nmFile <> vbNullString Then
   'DELETING THE PREVIOUS REPORT IF IT EXISTS
   'ExistFile = Dir(App.Path & "\" & nmFile)
   'If ExistFile <> "" Then Kill (App.Path & "\" & nmFile)
   On Error GoTo errHND
   ExistFile = dir(cmDlg.FileName)
   If ExistFile <> "" Then Kill (cmDlg.FileName)
   Export2XL 1, App.Path & "\dinkes07.MDB", nmTabel, nmFile, cmDlg.FileName, frmExim
  End If
 
 Case 0     'Import Data
  Dim xlBook As Excel.Workbook, irow As Integer
  Dim sMySql As String, Sql2 As String
  Dim no_trans, bulan, tahun, kdpuskesmas, jumlahtt, pelapor
  Dim kdPenyakit, usial1, usiap1, total1
  Dim usial2, usiap2, total2, usial3, usiap3, total3
  Dim usial4, usiap4, total4, usial5, usiap5, total5
  Dim usial6, usiap6, total6, usial7, usiap7, total7
  Dim usial8, usiap8, total8, usial9, usiap9, total9
  Dim usial10, usiap10, total10, usial11, usiap11, total11
  Dim usial12, usiap12, total12, totall, totalp
  Dim kdsubkegiatan, jumlahl, jumlahp, keterangan, total
    

  Select Case cmbData.ListIndex
  Case 0
   cmDlg.Filter = "(*.xls)|*.xls"
   cmDlg.ShowOpen
   nmFile = cmDlg.FileName
  Case 1
   cmDlg.Filter = "(*.xls)|*.xls"
   cmDlg.ShowOpen
   nmFile = cmDlg.FileName
  Case 2
   cmDlg.Filter = "(*.xls)|*.xls"
   cmDlg.ShowOpen
   nmFile = cmDlg.FileName
  End Select
  
  'On Error GoTo errHND
  If nmFile <> vbNullString Then
   Set xlBook = GetObject(nmFile)
   
   irow = 2
   pgb.Min = 0
   pgb.Value = 0
   While xlBook.Worksheets(1).Cells(irow, 1).Value <> vbNullString
   
    no_trans = xlBook.Worksheets(1).Cells(irow, 1).Value
    bulan = xlBook.Worksheets(1).Cells(irow, 2).Value
    tahun = xlBook.Worksheets(1).Cells(irow, 3).Value
    kdpuskesmas = xlBook.Worksheets(1).Cells(irow, 4).Value
    jumlahtt = xlBook.Worksheets(1).Cells(irow, 5).Value
    pelapor = xlBook.Worksheets(1).Cells(irow, 6).Value
    
    Select Case cmbData.ListIndex
    Case 0
     sMySql = "insert into tbtranspenyakit values('" & no_trans & _
              "'," & bulan & "," & tahun & ",'" & kdpuskesmas & _
              "'," & jumlahtt & ",'" & pelapor & "')"
      
     kdPenyakit = xlBook.Worksheets(1).Cells(irow, 7).Value
     usial1 = xlBook.Worksheets(1).Cells(irow, 8).Value
     usiap1 = xlBook.Worksheets(1).Cells(irow, 9).Value
     total1 = xlBook.Worksheets(1).Cells(irow, 10).Value
     usial2 = xlBook.Worksheets(1).Cells(irow, 11).Value
     usiap2 = xlBook.Worksheets(1).Cells(irow, 12).Value
     total2 = xlBook.Worksheets(1).Cells(irow, 13).Value
     usial3 = xlBook.Worksheets(1).Cells(irow, 14).Value
     usiap3 = xlBook.Worksheets(1).Cells(irow, 15).Value
     total3 = xlBook.Worksheets(1).Cells(irow, 16).Value
     usial4 = xlBook.Worksheets(1).Cells(irow, 17).Value
     usiap4 = xlBook.Worksheets(1).Cells(irow, 18).Value
     total4 = xlBook.Worksheets(1).Cells(irow, 19).Value
     usial5 = xlBook.Worksheets(1).Cells(irow, 20).Value
     usiap5 = xlBook.Worksheets(1).Cells(irow, 21).Value
     total5 = xlBook.Worksheets(1).Cells(irow, 22).Value
     usial6 = xlBook.Worksheets(1).Cells(irow, 23).Value
     usiap6 = xlBook.Worksheets(1).Cells(irow, 24).Value
     total6 = xlBook.Worksheets(1).Cells(irow, 25).Value
     usial7 = xlBook.Worksheets(1).Cells(irow, 26).Value
     usiap7 = xlBook.Worksheets(1).Cells(irow, 27).Value
     total7 = xlBook.Worksheets(1).Cells(irow, 28).Value
     usial8 = xlBook.Worksheets(1).Cells(irow, 29).Value
     usiap8 = xlBook.Worksheets(1).Cells(irow, 30).Value
     total8 = xlBook.Worksheets(1).Cells(irow, 31).Value
     usial9 = xlBook.Worksheets(1).Cells(irow, 32).Value
     usiap9 = xlBook.Worksheets(1).Cells(irow, 33).Value
     total9 = xlBook.Worksheets(1).Cells(irow, 34).Value
     usial10 = xlBook.Worksheets(1).Cells(irow, 35).Value
     usiap10 = xlBook.Worksheets(1).Cells(irow, 36).Value
     total10 = xlBook.Worksheets(1).Cells(irow, 37).Value
     usial11 = xlBook.Worksheets(1).Cells(irow, 38).Value
     usiap11 = xlBook.Worksheets(1).Cells(irow, 39).Value
     total11 = xlBook.Worksheets(1).Cells(irow, 40).Value
     usial12 = xlBook.Worksheets(1).Cells(irow, 41).Value
     usiap12 = xlBook.Worksheets(1).Cells(irow, 42).Value
     total12 = xlBook.Worksheets(1).Cells(irow, 43).Value
     totall = xlBook.Worksheets(1).Cells(irow, 44).Value
     totalp = xlBook.Worksheets(1).Cells(irow, 45).Value
     
     Sql2 = "insert into tbTransDtlPenyakit values('" & _
       no_trans & "','" & kdPenyakit & "'," & usial1 & _
       "," & usiap1 & "," & total1 & "," & usial2 & _
       "," & usiap2 & "," & total2 & "," & usial3 & _
       "," & usiap3 & "," & total3 & "," & usial4 & _
       "," & usiap4 & "," & total4 & "," & usial5 & _
       "," & usiap5 & "," & total5 & "," & usial6 & _
       "," & usiap6 & "," & total6 & "," & usial7 & _
       "," & usiap7 & "," & total7 & "," & usial8 & _
       "," & usiap8 & "," & total8 & "," & usial9 & _
       "," & usiap9 & "," & total9 & "," & usial10 & _
       "," & usiap10 & "," & total10 & "," & usial11 & _
       "," & usiap11 & "," & total11 & "," & usial12 & _
       "," & usiap12 & "," & total12 & ",''," & _
       totall & "," & totalp & ")"
    
    Case 1, 2
     kdsubkegiatan = xlBook.Worksheets(1).Cells(irow, 7).Value
     jumlahl = xlBook.Worksheets(1).Cells(irow, 8).Value
     jumlahp = xlBook.Worksheets(1).Cells(irow, 9).Value
     keterangan = xlBook.Worksheets(1).Cells(irow, 10).Value
     total = xlBook.Worksheets(1).Cells(irow, 11).Value
    
     If cmbData.ListIndex = 1 Then
      sMySql = "insert into tbtransGK values('" & no_trans & _
       "'," & bulan & "," & tahun & ",'" & kdpuskesmas & "'," & jumlahtt & _
       ",'" & pelapor & "')"
              
      Sql2 = "insert into tbtransdtlGK values('" & no_trans & _
       "','" & kdsubkegiatan & "'," & jumlahl & "," & _
       jumlahp & ",'" & keterangan & "'," & total & ")"
     Else
      sMySql = "insert into tbtranskegiatan values('" & no_trans & _
       "'," & bulan & "," & tahun & ",'" & kdpuskesmas & "'," & jumlahtt & _
       ",'" & pelapor & "')"
     
      Sql2 = "insert into tbtransdtlKegiatan values('" & no_trans & _
       "','" & kdsubkegiatan & "'," & jumlahl & "," & _
       jumlahp & ",'" & keterangan & "'," & total & ")"
     End If
    End Select
    
    'Save to Database
    con.BeginTrans
    
    On Error Resume Next
    con.Execute sMySql
    
    con.Execute Sql2
    con.CommitTrans
        
    irow = irow + 1
    pgb.Max = irow - 2
    pgb.Value = pgb.Max
   Wend
   MsgBox "Proses Upload Data Selesai", vbOKOnly + vbInformation, "DINAS KESEHATAN"
   Set xlBook = Nothing
  End If
 End Select
errHND:
 'error handler
End Sub

Private Sub Form_Load()
 Set con = New ADODB.Connection
 con.CursorLocation = adUseClient
 'con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
 '            "Data Source=" & App.Path & "\dinkes07.mdb;"
 con.Open "DSN=dinkesLab"
 cmbData.ListIndex = 0
 
 cmbBulan.ListIndex = Month(Now) - 1
 txtTahun.Text = Year(Now)
End Sub
