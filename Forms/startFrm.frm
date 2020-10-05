VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm startFrm 
   BackColor       =   &H8000000F&
   Caption         =   "Sistem Informasi SP2TP (LB1, LB3 Dan LB4)"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   -2175
   ClientWidth     =   11880
   Icon            =   "startFrm.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "startFrm.frx":0ECA
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Penyakit"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Kegiatan Bulanan"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Gizi & KIA"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Transaksi LB1"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Transaksi LB3"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Transaksi LB4"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Maintenance"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Export/Import Data"
            ImageIndex      =   17
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Book report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Member report"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Issue report"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Laporan Transaksi"
            ImageIndex      =   10
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "LB1 - Kunjungan Pasien"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "LB3 - Gizi & KIA"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "LB4 - Kegiatan Bulanan"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Keluar"
            ImageIndex      =   11
         EndProperty
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "startFrm.frx":C155A
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   7920
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistem SP2TP Untuk Dinas Kesehatan "
            TextSave        =   "Sistem SP2TP Untuk Dinas Kesehatan "
            Object.ToolTipText     =   "Programmer By Hidayat Widhi, Editing By Habibur Rahman"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   970
            MinWidth        =   970
            Text            =   "   User"
            TextSave        =   "   User"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1111
            MinWidth        =   882
            Text            =   "  Today"
            TextSave        =   "  Today"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2293
            MinWidth        =   2293
            TextSave        =   "17/02/2009"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1:28 AM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3863
            Text            =   "Contact : widhi79@gmail.com"
            TextSave        =   "Contact : widhi79@gmail.com"
            Object.ToolTipText     =   "Created by : DHS Team 2007"
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "startFrm.frx":C16BC
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C181E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C24F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C31D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C3EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C4B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C5860
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C653A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C7214
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C7EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C8BC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":C98A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":CA57C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":CB256
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":CBF30
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":CCC0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":CD8E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "startFrm.frx":CE5BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_database 
      Caption         =   "&Database"
      Begin VB.Menu sm_database 
         Caption         =   "P&ropinsi"
         Index           =   0
         Shortcut        =   ^R
      End
      Begin VB.Menu sm_database 
         Caption         =   "K&abupaten"
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu sm_database 
         Caption         =   "&Kecamatan"
         Index           =   2
         Shortcut        =   ^K
      End
      Begin VB.Menu sm_database 
         Caption         =   "&Puskesmas"
         Index           =   3
         Shortcut        =   ^P
      End
      Begin VB.Menu sm_database 
         Caption         =   "P&enyakit (LB1)"
         Index           =   4
         Begin VB.Menu subsm_penyakit 
            Caption         =   "&Jenis Penyakit"
            Index           =   0
            Shortcut        =   ^J
         End
         Begin VB.Menu subsm_penyakit 
            Caption         =   "&Daftar Penyakit"
            Index           =   1
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu sm_database 
         Caption         =   "Ke&giatan (LB3 - LB4)"
         Index           =   5
         Begin VB.Menu subsm_kegiatan 
            Caption         =   "Jeni&s Kegiatan - (LB4 - LB3)"
            Index           =   0
            Shortcut        =   ^S
         End
         Begin VB.Menu subsm_kegiatan 
            Caption         =   "Kegiata&n Puskesmas (LB4)"
            Index           =   1
            Shortcut        =   ^N
         End
         Begin VB.Menu subsm_kegiatan 
            Caption         =   "Gi&zi dan KIA (LB3)"
            Index           =   2
            Shortcut        =   ^Z
         End
      End
      Begin VB.Menu sm_database 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu sm_database 
         Caption         =   "&Logoff"
         Index           =   7
         Shortcut        =   ^L
      End
      Begin VB.Menu sm_database 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu sm_database 
         Caption         =   "K&eluar"
         Index           =   9
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnu_transaksi 
      Caption         =   "T&ransaksi"
      Begin VB.Menu sm_transaksi 
         Caption         =   "LB1 - Kunjungan Pas&ien"
         Index           =   0
         Shortcut        =   ^I
      End
      Begin VB.Menu sm_transaksi 
         Caption         =   "LB4 - Kegiatan Puskes&mas"
         Index           =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu sm_transaksi 
         Caption         =   "LB3 - Data &Gizi dan KIA"
         Index           =   2
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnu_administer 
      Caption         =   "&Tools"
      Begin VB.Menu sm_admin 
         Caption         =   "Maintenan&ce Transaksi"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu sm_admin 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu sm_admin 
         Caption         =   "E&xim Excel"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu sm_admin 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu sm_admin 
         Caption         =   "Back&up Data"
         Index           =   4
         Shortcut        =   ^U
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Mnu_rep 
      Caption         =   "&Laporan"
      Begin VB.Menu sm_Laporan 
         Caption         =   "LB1 - Kunjungan Pasien"
         Index           =   0
      End
      Begin VB.Menu sm_Laporan 
         Caption         =   "LB3 - Gizi dan KIA"
         Index           =   1
      End
      Begin VB.Menu sm_Laporan 
         Caption         =   "LB4 - Kegiatan Bulanan"
         Index           =   2
      End
      Begin VB.Menu sm_Laporan 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu sm_Laporan 
         Caption         =   "10 Besar Penyakit"
         Index           =   4
         Begin VB.Menu sm_10besar 
            Caption         =   "Per Tahun/Bulan (Kab.Gresik)"
            Index           =   0
         End
         Begin VB.Menu sm_10besar 
            Caption         =   "Per TRIWULAN Kab. Gresik"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu sm_10besar 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu sm_10besar 
            Caption         =   "Per Tahun/Bulan (Puskesmas)"
            Index           =   3
         End
      End
      Begin VB.Menu sm_Laporan 
         Caption         =   "Rekap &LB"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_win 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnu_help 
      Caption         =   "&Help"
      Begin VB.Menu sm_help 
         Caption         =   "Panduan Aplikasi "
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu sm_help 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu sm_help 
         Caption         =   "Tentang Program"
         Index           =   2
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "startFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim slogoff As Boolean

Private Sub MDIForm_Activate()
 'Picture1.Width = startFrm.Width
 'Image1.Width = startFrm.Width
 'Image1.Height = Picture1.Height
End Sub

Private Sub MDIForm_Load()
 Select Case appTipe
 Case 0     'Dinkes
  sm_database.Item(0).Visible = False
  sm_database.Item(1).Visible = False
 Case 1     'Puskesmas
  sm_database.Item(0).Visible = False
  sm_database.Item(1).Visible = False
  sm_database.Item(2).Visible = False
  sm_database.Item(3).Visible = False
 End Select
 slogoff = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 If Not slogoff Then
  If MsgBox("Anda Yakin Akan Keluar ?", vbExclamation + vbOKCancel, "DINKES - SP2TP") = vbOK Then
   Unload frmPropinsi
   Unload frmKabupaten
   Unload frmKecamatan
   Unload frmPuskesmas
   Unload frmJenisPenyakit
   Unload frmPenyakit
   Unload frmJenisKegiatan
   Unload frmKegBulanan
   Unload frmGiziKia
   Unload frmNTransGK
   Unload frmNTransKegiatan
   Unload frmNTransPenyakit
   Unload frmMaintenance
   Unload frmExim
   Unload rptLB1
   Unload rptLB3
   Unload rptLB4
   Unload rptForm
   Unload frmLogin
  Else
   Cancel = True
  End If
 End If
End Sub

Private Sub sm_10besar_Click(Index As Integer)
 Select Case Index
 Case 0
  Load rpt10BlThKab
  rpt10BlThKab.Show
 Case 1
  Load rpt103BlThKab
  rpt103BlThKab.Show
 Case 3
  Load rpt10BlThKabPus
  rpt10BlThKabPus.Show
 End Select
End Sub

Private Sub sm_admin_Click(Index As Integer)
 Select Case Index
 Case 0
  Load frmMaintenance
  frmMaintenance.Show
 Case 2
  Load frmExim
  frmExim.Show
 Case 4
 End Select
End Sub

'Private Sub MDIForm_Load()
'Me.Show
'Me.Enabled = False
  'setting toolbar images
'With Toolbar2
'Set .ImageList = ImageList1
'.Buttons(2).Image = 1
'.Buttons(3).Image = 7
'.Buttons(5).Image = 5
'.Buttons(6).Image = 6
'.Buttons(7).Image = 14
'.Buttons(8).Image = 2
'.Buttons(9).Image = 3
'.Buttons(11).Image = 10
'.Buttons(13).Image = 8
'.Buttons(14).Image = 9
'.Buttons(16).Image = 12
'.Buttons(17).Image = 13
'.Buttons(19).Image = 4
'.Buttons(20).Image = 11
'End With
'sbStatusBar.Panels(3).Text = "Login"
'End Sub


'Private Sub sbStatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
'ShellExecute Me.hWnd, vbNullString, "DHS Team - (031)60607774", vbNullString, vbNullString, SW_SHOWNORMAL
'End Sub
'Private Sub sm_about_Click()
'Load frmAbout
'frmAbout.Show
'End Sub
'Private Sub sm_backup_Click()
'Load Frm_backup
'Frm_backup.Show
'End Sub

'Private Sub sm_bookrpt_Click()
'Load Frm_bookrpt
'Frm_bookrpt.Show
'End Sub

'Private Sub sm_books_Click()
'Load Frm_books
'Frm_books.Show
'End Sub

'Private Sub sm_calculator_Click()
'On Error GoTo errHandle
'    Dim a As Double
'    a = Shell("C:\WINDOWS\System32\calc.exe", vbNormalFocus)
'    Exit Sub
'errHandle:
'    MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
'    Resume Next
'End Sub

'Private Sub sm_employees_Click()
'Load Frm_Employees
'Frm_Employees.Show
'End Sub

'Private Sub sm_exit_Click()
'Unload Me
'End Sub

'Private Sub sm_fine_Click()
'Load Frm_Fine
'Frm_Fine.Show
'End Sub

'Private Sub sm_global_Click()
'Load Frm_global
'Frm_global.Show
'End Sub

'Private Sub sm_help_Click()
' Dim nRet As Integer
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
'End Sub

'Private Sub sm_hsearch_Click()
'    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        On Error Resume Next
'        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If
'End Sub

'Private Sub sm_issret_Click()
'Load Frm_issretrpt
'Frm_issretrpt.Show
'End Sub

'Private Sub sm_issue_Click()
'Load Frm_issue
'Frm_issue.Show
'End Sub

'Private Sub sm_logoff_Click()
'If MsgBox("Are You Sure you want to logoff ?", vbExclamation + vbOKCancel, "Library Management System") = vbOK Then
'Call logoff
'DoEvents
'End If
'End Sub

'Private Sub sm_member_Click()
'Load Frm_memrpt
'Frm_memrpt.Show
'End Sub

'Private Sub sm_members_Click()
'Load Frm_members
'Frm_members.Show
'End Sub

'Private Sub sm_notepad_Click()
'On Error GoTo errcode
'    Dim a As Double
'    a = Shell("C:\WINDOWS\System32\notepad.exe", vbNormalFocus)
'    Exit Sub
'errcode:
'    MsgBox "Unable to run Notepad Utility on your computer", vbInformation, "Error in opening!!!"
'    Resume Next
'End Sub

'Private Sub sm_return_Click()
'Load Frm_return
'Frm_return.Show
'End Sub

'Private Sub sm_search_Click()
'Load Frm_search
'Frm_search.Show
'End Sub

'Private Sub sm_settings_Click()
'Load Frm_settings
'Frm_settings.Show
'End Sub

'Private Sub smnu_keyboard_Click()
'Load Frm_keyboard
'Frm_keyboard.Show
'End Sub

'Private Sub Toolbar2_ButtonClick(ByVal button As MSComctlLib.button)
'Select Case button.Index
'    Case 2: Call sm_books_Click
'    Case 3: Call sm_members_Click
'    Case 5: Call sm_issue_Click
'    Case 6: Call sm_return_Click
'    Case 7: Call sm_fine_Click
'    Case 8: Call sm_search_Click
'    Case 9: Call sm_global_Click
'    Case 11: 'add report
'    Case 13: Call sm_calculator_Click
'    Case 14: Call sm_notepad_Click
'    Case 16: Call smnu_keyboard_Click
'    Case 17: Call sm_about_Click
'    Case 19: Call sm_logoff_Click
'    Case 20: Call sm_exit_Click
'End Select
'End Sub

'Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'Select Case ButtonMenu.Index
'    Case 1:
'         Call sm_bookrpt_Click
'    Case 2:
'         Call sm_member_Click
'    Case 3:
'         Call sm_issret_Click
'End Select
'End Sub

Private Sub sm_database_Click(Index As Integer)
 Select Case Index
 Case 0     'Database propinsi
  Load frmPropinsi
  frmPropinsi.Show
 Case 1     'Database kabupaten
  Load frmKabupaten
  frmKabupaten.Show
 Case 2     'Database kecamatan
  Load frmKecamatan
  frmKecamatan.Show
 Case 3     'Database puskesmas
  Load frmPuskesmas
  frmPuskesmas.Show
 Case 7     'Log Off
  If MsgBox("Anda Yakin Ingin logoff ?", vbExclamation + vbOKCancel, "DINKES - SP2TP") = vbOK Then
   'Call logoff
   'DoEvents
   slogoff = True
   Unload frmPropinsi
   Unload frmKabupaten
   Unload frmKecamatan
   Unload frmPuskesmas
   Unload frmJenisPenyakit
   Unload frmPenyakit
   Unload frmJenisKegiatan
   Unload frmKegBulanan
   Unload frmGiziKia
   Unload frmTransGK
   Unload frmTransKegiatan
   Unload frmTransPenyakit
   Unload frmMaintenance
   Unload frmExim
   Unload rptLB1
   Unload rptLB3
   Unload rptLB4
   Unload rptForm
   Unload Me
   frmLogin.Show
   slogoff = False
  End If
 Case 9     'Keluar
  Unload Me
 End Select
End Sub

Private Sub sm_laporan_Click(Index As Integer)
 Select Case Index
 Case 0
  Load rptLB1
  rptLB1.Show
 Case 1
  Load rptLB3
  rptLB3.Show
 Case 2
  Load rptLB4
  rptLB4.Show
 Case 5
  Load rptRekapLB
  rptRekapLB.Show
 End Select
End Sub

Private Sub sm_transaksi_Click(Index As Integer)
 Select Case Index
 Case 0
  Load frmNTransPenyakit
  frmNTransPenyakit.Show
 Case 1
  Load frmNTransKegiatan
  frmNTransKegiatan.Show
 Case 2
  Load frmNTransGK
  frmNTransGK.Show
 End Select
End Sub

Private Sub subsm_kegiatan_Click(Index As Integer)
 Select Case Index
 Case 0
  Load frmJenisKegiatan
  frmJenisKegiatan.Show
 Case 1
  Load frmKegBulanan
  frmKegBulanan.Show
 Case 2
  Load frmGiziKia
  frmGiziKia.Show
 End Select
End Sub

Private Sub subsm_penyakit_Click(Index As Integer)
 Select Case Index
 Case 0
  Load frmJenisPenyakit
  frmJenisPenyakit.Show
 Case 1
  Load frmPenyakit
  frmPenyakit.Show
 End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Index
   Case 2: Call subsm_penyakit_Click(1)
   Case 3: Call subsm_kegiatan_Click(1)
   Case 4: Call subsm_kegiatan_Click(2)
   Case 6: Call sm_transaksi_Click(0)
   Case 7: Call sm_transaksi_Click(2)
   Case 8: Call sm_transaksi_Click(1)
   Case 10: Call sm_admin_Click(0)
   Case 11: Call sm_admin_Click(2)
   Case 14: Call sm_database_Click(7)
   Case 15: Call sm_database_Click(9)
 End Select
End Sub

Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Select Case ButtonMenu.Index
  Case 1:
     Call sm_laporan_Click(0)
  Case 2:
     Call sm_laporan_Click(1)
  Case 3:
     Call sm_laporan_Click(2)
  End Select
End Sub
