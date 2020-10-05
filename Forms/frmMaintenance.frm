VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaintenance 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
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
      Left            =   9840
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.ComboBox cmbData 
      Height          =   315
      ItemData        =   "frmMaintenance.frx":0000
      Left            =   1560
      List            =   "frmMaintenance.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin MSComctlLib.ListView lstData 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Klik Dobel Untuk Edit || Del untuk Hapus Data"
      Top             =   600
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No Transaksi"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kode Puskesmas"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Puskesmas"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Jumlah T.T"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Jumlah Pelapor"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Bulan"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tahun"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Transaksi"
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
      Left            =   120
      TabIndex        =   2
      Top             =   255
      Width           =   1455
   End
End
Attribute VB_Name = "frmMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Private myList1 As ListItem
Private Status As String
Dim lstSql As String, lstPil As Byte
Dim nChecked As Integer
Dim nBulan As Variant, nTahun As Variant

Option Explicit

Private Sub RefreshList(Pil As Byte)
 Dim i As Integer
 Dim rsPerusahaan As New ADODB.Recordset

 lstPil = Pil
 Select Case Pil
 Case 1
  lstSql = "SELECT * from qDistincTransPus"
 Case 2
  lstSql = "SELECT * from qDistincTransPusG"
 Case 3
  lstSql = "SELECT * from qDistincTransPusK"
 End Select
 Set rsPerusahaan = con.Execute(lstSql)
 If Not rsPerusahaan.EOF Then
  rsPerusahaan.MoveFirst
  lstData.ListItems.Clear
  While Not rsPerusahaan.EOF
   Set myList1 = lstData.ListItems.Add(, , rsPerusahaan.Fields(0).Value)
   myList1.SubItems(1) = rsPerusahaan.Fields(1).Value
   myList1.SubItems(2) = rsPerusahaan.Fields(2).Value
   myList1.SubItems(3) = rsPerusahaan.Fields(3).Value
   myList1.SubItems(4) = rsPerusahaan.Fields(4).Value
   myList1.SubItems(5) = rsPerusahaan.Fields(5).Value
   myList1.SubItems(6) = rsPerusahaan.Fields(6).Value
   rsPerusahaan.MoveNext
  Wend
 Else
  lstData.ListItems.Clear
 End If
 rsPerusahaan.Close
 Set rsPerusahaan = Nothing
End Sub

Private Sub cmbData_Click()
 Select Case cmbData.ListIndex
 Case 0
  RefreshList 1
 Case 1
  RefreshList 2
 Case 2
  RefreshList 3
 End Select
End Sub

Private Sub cmdBtn_Click()
 Unload Me
End Sub

Private Sub Form_Activate()
 Select Case cmbData.ListIndex
 Case 0
  RefreshList 1
 Case 1
  RefreshList 2
 Case 2
  RefreshList 3
 End Select
End Sub

Private Sub Form_Load()
 Set con = New ADODB.Connection
 con.CursorLocation = adUseClient
 'con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
 '            "Data Source=" & App.Path & "\dinkes07.mdb;"
 con.Open "DSN=dinkesLab"
End Sub

Private Sub Form_Unload(Cancel As Integer)
 con.Close
End Sub

Private Sub lstData_DblClick()
 Dim strJenis As String
 Dim i As Byte
 Dim rsFind As Recordset
 
 'On Error GoTo errHND
 
 Status = "Edit"
 Select Case lstPil
 Case 1
  With frmNTransPenyakit
   .noTrans = lstData.ListItems(lstData.SelectedItem.Index).Text
   .newBtn.Enabled = False
   .cmbData(0).ListIndex = val(lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(5).Text) - 1
   .cmbData(0).Locked = True
   .txtdata(0).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(6).Text
   .txtdata(0).Locked = True
   .txtdata(1).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(1).Text
   .txtdata(1).Enabled = False
   .txtdata(2).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(2).Text
   .txtdata(3).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(3).Text
   .txtdata(3).Enabled = True
   .txtdata(4).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(4).Text
   .txtdata(4).Enabled = True
   
   .cmbData(1).Enabled = True
   .cmbData(1).ListIndex = 0
   .lstData.Enabled = True
   .lstData.ListIndex = 0
   For i = 0 To 38
    .pvNum(i).Enabled = True
   Next
   .Show
  End With
 Case 2
  With frmNTransGK
   .noTrans = lstData.ListItems(lstData.SelectedItem.Index).Text
   .cmbData(0).ListIndex = val(lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(5).Text) - 1
   .cmbData(0).Locked = True
   .txtdata(0).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(6).Text
   .txtdata(0).Locked = True
   .txtdata(1).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(1).Text
   .txtdata(1).Enabled = False
   .txtdata(2).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(2).Text
   .txtdata(3).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(3).Text
   .txtdata(3).Enabled = True
   .txtdata(4).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(4).Text
   .txtdata(4).Enabled = True
  
   .cmbData(1).Enabled = True
   .cmbData(1).ListIndex = 0
   .lstData.Enabled = True
   .lstData.ListIndex = 0
   For i = 0 To 2
     .pvNum(i).Enabled = True
   Next
   .Show
  End With
 Case 3
  With frmNTransKegiatan
   .noTrans = lstData.ListItems(lstData.SelectedItem.Index).Text
   .cmbData(0).ListIndex = val(lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(5).Text) - 1
   .cmbData(0).Locked = True
   .txtdata(0).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(6).Text
   .txtdata(0).Locked = True
   .txtdata(1).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(1).Text
   .txtdata(1).Enabled = False
   .txtdata(2).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(2).Text
   .txtdata(3).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(3).Text
   .txtdata(3).Enabled = True
   .txtdata(4).Text = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(4).Text
   .txtdata(4).Enabled = True
  
   .cmbData(1).Enabled = True
   .cmbData(1).ListIndex = 0
   .lstData.Enabled = True
   .lstData.ListIndex = 0
    For i = 0 To 2
    .pvNum(i).Enabled = True
    Next
   .Show
  End With
 End Select
 Status = "Non-Edit"
 'Exit Sub
 
'errHND:
' MsgBox "Tidak Ada Data Yang Dipilih", vbOKOnly + vbInformation, "Dinas Kesehatan"
End Sub

Private Sub lstData_ItemCheck(ByVal Item As MSComctlLib.ListItem)
 Item.Selected = True
 If Item.Checked Then
  nChecked = nChecked + 1
 Else
  nChecked = nChecked - 1
 End If
End Sub

Private Sub lstData_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim Ask As Variant, i As Integer
 Dim kdPus As Variant, noTrans As String
 
 If KeyCode = vbKeyDelete And _
    nChecked > 0 And Status <> "Edit" Then
  Ask = MsgBox("Apakah Data Akan Dihapus ?", vbYesNo + vbQuestion + vbDefaultButton2, "SIT")
  If Ask = vbYes Then
   If nChecked = 1 Then
    noTrans = lstData.ListItems(lstData.SelectedItem.Index).Text
    kdPus = lstData.ListItems(lstData.SelectedItem.Index).ListSubItems(1).Text
    Select Case lstPil
    Case 1
     lstSql = "DELETE FROM tbTransPenyakit " & _
          "where no_trans='" & noTrans & _
          "' and kdPuskesmas='" & kdPus & "'"
    Case 2
     lstSql = "DELETE FROM tbTransGK " & _
          "where no_trans='" & noTrans & _
          "' and kdPuskesmas='" & kdPus & "'"
    Case 3
     lstSql = "DELETE FROM tbTransKegiatan " & _
          "where no_trans='" & noTrans & _
          "' and kdPuskesmas='" & kdPus & "'"
    End Select
    con.BeginTrans
    con.Execute lstSql
    con.CommitTrans
   Else
    For i = 1 To lstData.ListItems.count
     If lstData.ListItems(i).Checked Then
      noTrans = lstData.ListItems(i).Text
      kdPus = lstData.ListItems(i).ListSubItems(1).Text
      Select Case lstPil
      Case 1
       lstSql = "DELETE FROM tbTransPenyakit " & _
          "where no_trans='" & noTrans & _
          "' and kdPuskesmas='" & kdPus & "'"
      Case 2
       lstSql = "DELETE FROM tbTransGK " & _
          "where no_trans='" & noTrans & _
          "' and kdPuskesmas='" & kdPus & "'"
      Case 3
       lstSql = "DELETE FROM tbTransKegiatan " & _
          "where no_trans='" & noTrans & _
          "' and kdPuskesmas='" & kdPus & "'"
      End Select
      con.BeginTrans
      con.Execute lstSql
      con.CommitTrans
     End If
    Next
   End If
   RefreshList lstPil
  End If
 Else
  If KeyCode = vbKeyDelete Then _
   MsgBox "Pilih Data Yang Akan Dihapus", vbOKOnly + vbInformation, "Dinas Kesehatan"
 End If
End Sub
