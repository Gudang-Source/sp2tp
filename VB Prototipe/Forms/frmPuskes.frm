VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPuskesmas 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Puskesmas"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   8610
   Begin MSComctlLib.StatusBar frmStbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   5340
      Width           =   8610
      _ExtentX        =   15187
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6271
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6271
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstData 
      Height          =   5295
      Left            =   2520
      TabIndex        =   10
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   16744576
      TabCaption(0)   =   "View Data"
      TabPicture(0)   =   "frmPuskes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DataGrid"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtdata(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdButton(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbData"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Form Isian"
      TabPicture(1)   =   "frmPuskes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtdata(12)"
      Tab(1).Control(1)=   "txtdata(11)"
      Tab(1).Control(2)=   "txtdata(10)"
      Tab(1).Control(3)=   "txtdata(9)"
      Tab(1).Control(4)=   "txtdata(8)"
      Tab(1).Control(5)=   "txtdata(7)"
      Tab(1).Control(6)=   "txtdata(6)"
      Tab(1).Control(7)=   "txtdata(5)"
      Tab(1).Control(8)=   "txtdata(4)"
      Tab(1).Control(9)=   "txtdata(3)"
      Tab(1).Control(10)=   "txtdata(2)"
      Tab(1).Control(11)=   "txtdata(1)"
      Tab(1).Control(12)=   "Label11"
      Tab(1).Control(13)=   "Label9"
      Tab(1).Control(14)=   "Label8"
      Tab(1).Control(15)=   "Label7"
      Tab(1).Control(16)=   "Label6"
      Tab(1).Control(17)=   "Label5"
      Tab(1).Control(18)=   "Label4"
      Tab(1).Control(19)=   "Line3"
      Tab(1).Control(20)=   "Line1"
      Tab(1).Control(21)=   "Label2"
      Tab(1).Control(22)=   "Label1"
      Tab(1).ControlCount=   23
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   12
         Left            =   -73200
         TabIndex        =   26
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   11
         Left            =   -73200
         TabIndex        =   25
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   10
         Left            =   -73200
         TabIndex        =   24
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   1020
         Index           =   9
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   8
         Left            =   -72120
         TabIndex        =   22
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   7
         Left            =   -73200
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   -72120
         TabIndex        =   20
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   -73200
         TabIndex        =   19
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   -72120
         TabIndex        =   18
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   -73200
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4440
         Width           =   3735
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   -73200
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Cari"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   4920
         MouseIcon       =   "frmPuskes.frx":0038
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Search"
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   4800
         Width           =   4695
      End
      Begin VB.TextBox txtdata 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   -73200
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid 
         Height          =   3735
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   -74880
         TabIndex        =   41
         Top             =   4440
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   -74880
         TabIndex        =   40
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Telepon"
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
         Left            =   -74880
         TabIndex        =   39
         Top             =   3720
         Width           =   825
      End
      Begin VB.Label Label7 
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kecamatan"
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
         Left            =   -74880
         TabIndex        =   37
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kabupaten"
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
         Left            =   -74880
         TabIndex        =   36
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Propinsi"
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
         Left            =   -74880
         TabIndex        =   35
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "FILTER"
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
         Left            =   120
         TabIndex        =   34
         Top             =   4440
         Width           =   825
      End
      Begin VB.Line Line3 
         X1              =   -74880
         X2              =   -69120
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line1 
         X1              =   -74880
         X2              =   -69120
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nama Puskesmas"
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
         Left            =   -74850
         TabIndex        =   30
         Top             =   1200
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode Puskesmas"
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
         Left            =   -74850
         TabIndex        =   28
         Top             =   840
         Width           =   1635
      End
   End
   Begin VB.Frame frmButton 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   27
      Top             =   -120
      Width           =   2415
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   9
         Left            =   1920
         MouseIcon       =   "frmPuskes.frx":0793
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":08E5
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move Last"
         Top             =   4320
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   8
         Left            =   1560
         MouseIcon       =   "frmPuskes.frx":0B37
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":0C89
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Move Next"
         Top             =   4320
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   7
         Left            =   480
         MouseIcon       =   "frmPuskes.frx":0E95
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":0FE7
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move Previous"
         Top             =   4320
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   6
         Left            =   120
         MouseIcon       =   "frmPuskes.frx":11F6
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":1348
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Move First"
         Top             =   4320
         UseMaskColor    =   -1  'True
         Width           =   345
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
         Height          =   615
         Index           =   5
         Left            =   120
         MouseIcon       =   "frmPuskes.frx":1597
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":16E9
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Save record"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   1200
         MouseIcon       =   "frmPuskes.frx":1C81
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":1DD3
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancel"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Caption         =   "T U T U P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   120
         MouseIcon       =   "frmPuskes.frx":2353
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":24A5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Caption         =   "Hapus Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   120
         MouseIcon       =   "frmPuskes.frx":2A1F
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":2B71
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Delete record"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Caption         =   "Edit Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   120
         MouseIcon       =   "frmPuskes.frx":30BB
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":320D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Edit record"
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdButton 
         Appearance      =   0  'Flat
         Caption         =   "Tambah Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         MaskColor       =   &H8000000F&
         MouseIcon       =   "frmPuskes.frx":37B2
         MousePointer    =   99  'Custom
         Picture         =   "frmPuskes.frx":3904
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Add new record"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl_total 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label lbl_rec 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   31
         Top             =   3960
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmPuskesmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoRec As ADODB.Recordset
Dim adoCon As ADODB.Connection
Dim tmpRec As ADODB.Recordset
Dim strSQL As String, saveFlag As Boolean
Dim exitFlag As Boolean, Editing As Boolean

Option Explicit

Private Sub cmdButton_Click(Index As Integer)
 Editing = False
 Select Case Index
 Case 0     'Tambah Data
  Editing = True
  If appTipe <> 1 Then
   sstData.Tab = 1
   'Make all entries in input mode enable
   Call SetLock(False)
   'clear the text field
   Call Clear
   saveFlag = True
   'lblStatus.Caption = " Add new record."
   Call SetButton(False)
  End If
 Case 1     'Edit Data
  Editing = True
  sstData.Tab = 1
  'Make all entries in input mode
   Call SetLock(False)
   saveFlag = False
  'message for status of mode
   'lblStatus.Caption = " Edit record"
   Call SetButton(False)
   If appTipe = 2 Then
    txtData(1).Enabled = False
    txtData(2).SetFocus
   End If
 Case 2     'Tutup Form
  exitFlag = True
  Unload Me
 Case 3     'Cari Data
  adoRec.MoveFirst
  adoRec.Find Trim(cmbData.Text) & " LIKE '%" & Trim(txtData(0).Text) & "%'"
 Case 4     'Hapus Data
  If appTipe <> 1 Then Call pHapus
 Case 5     'Simpan Record
  Call pSimpan
 Case 6     'First Record
  On Error GoTo GoNextError
  adoRec.MoveFirst
  'show thw current data record
 Case 7     'Previous Record
  On Error GoTo GoNextError
  If Not adoRec.BOF Then adoRec.MovePrevious
  If adoRec.BOF And adoRec.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoRec.MoveFirst
  End If
  'show thw current data record
 Case 8     'Next Record
  On Error GoTo GoNextError
  If Not adoRec.EOF Then adoRec.MoveNext
  If adoRec.EOF And adoRec.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoRec.MoveLast
  End If
  'show thw current data record
 Case 9     'Last Record
  On Error GoTo GoNextError
  adoRec.MoveLast
  'show thw current data record
 Case 10    'Batal
  Call SetLock(True)
  Call SetButton(True)
  If adoRec.BOF And adoRec.EOF Then
   Call Clear
   cmdButton(4).Enabled = False
   cmdButton(1).Enabled = False
   'GoTo newproc
  Else
   adoRec.MoveFirst
   Call Scatter
  End If
'newproc:
  txtData(1).Enabled = True
  If saveFlag Then
   txtData(1).SetFocus
  Else
   txtData(2).SetFocus
  End If
 End Select
 
 If (Index >= 6 And Index <= 9) _
    Or Index = 3 Then Call Scatter
 Exit Sub
 
GoNextError:
 MsgBox "Data Masih Kosong"
End Sub

Private Sub DataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 Call Scatter
End Sub

Private Sub Form_Load()
 Dim i As Integer
 
 'Image1.Picture = startFrm.ImageList1.ListImages(1).Picture
 Set adoCon = New ADODB.Connection
 adoCon.CursorLocation = adUseClient
 adoCon.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & App.Path & "\dinkes07.mdb;"
 strSQL = "select * from tbPuskesmas Order by kode"
 Set adoRec = New ADODB.Recordset
   adoRec.Open strSQL, adoCon, adOpenStatic, adLockOptimistic

 'show current record
  Call SetButton(True)
  Call Scatter
  Call txtActivate
  Set DataGrid.DataSource = adoRec
  DataGrid.ReBind
  
  cmbData.Clear
  For i = 0 To adoRec.Fields.Count - 1
   cmbData.AddItem adoRec.Fields(i).Name
  Next
  
  Call SetLock(True)
  exitFlag = False
End Sub

Private Sub Scatter()
 Call Locate
 If adoRec.EOF = False And adoRec.BOF = False Then
  txtData(1).Text = adoRec.Fields(0)
  txtData(2).Text = adoRec.Fields(1)
  txtData(3).Text = adoRec.Fields(2)
  txtData(5).Text = adoRec.Fields(3)
  txtData(7).Text = adoRec.Fields(4)
  txtData(9).Text = adoRec.Fields(5)
  txtData(10).Text = adoRec.Fields(6)
  txtData(11).Text = adoRec.Fields(7)
  txtData(12).Text = adoRec.Fields(8)
  
  Set tmpRec = New ADODB.Recordset
  strSQL = "select nama from tbProp where kode='" & _
           txtData(3).Text & "'"
  tmpRec.Open strSQL, adoCon, adOpenStatic, adLockOptimistic
  txtData(4).Text = tmpRec.Fields(0)
  tmpRec.Close
  
  strSQL = "select nama from tbKab where kode='" & _
           txtData(5).Text & "'"
  tmpRec.Open strSQL, adoCon, adOpenStatic, adLockOptimistic
  txtData(6).Text = tmpRec.Fields(0)
  tmpRec.Close
  
  strSQL = "select nama from tbKec where kode='" & _
           txtData(7).Text & "'"
  tmpRec.Open strSQL, adoCon, adOpenStatic, adLockOptimistic
  txtData(8).Text = tmpRec.Fields(0)
  tmpRec.Close
  
  cmdButton(4).Enabled = True
  cmdButton(1).Enabled = True
 Else
  cmdButton(4).Enabled = False
  cmdButton(1).Enabled = False
 End If
End Sub

Private Sub Locate()
 lbl_total.Caption = adoRec.RecordCount
 If adoRec.EOF = False And adoRec.BOF = False Then
  lbl_rec.Caption = adoRec.AbsolutePosition
 Else
  lbl_rec.Caption = 0
 End If
End Sub

Private Sub SetButton(val As Boolean)
  'enable buttons
  cmdButton(5).Enabled = Not val
  cmdButton(10).Enabled = Not val
  'disable buttons
  cmdButton(0).Enabled = val
  cmdButton(1).Enabled = val
  cmdButton(2).Enabled = val
  cmdButton(4).Enabled = val
  cmdButton(6).Enabled = val
  cmdButton(7).Enabled = val
  cmdButton(8).Enabled = val
  cmdButton(9).Enabled = val
End Sub

Private Sub SetLock(val As Boolean)
 Dim i As Integer
 
 For i = 1 To 12
  txtData(i).Locked = val
 Next
End Sub

Private Sub Clear()
 Dim i As Integer
 
 If appTipe = "0" Then
  txtData(1).Text = vbNullString
  txtData(2).Text = vbNullString
  For i = 7 To 12
   txtData(i).Text = vbNullString
  Next
 Else
  For i = 1 To 12
   txtData(i).Text = vbNullString
  Next
 End If

'set focus to fiRSt textbox
 txtData(1).SetFocus
End Sub

Private Sub pHapus()
  If MsgBox("Anda Yakin Record Ini Dihapus ?", vbYesNo + vbQuestion, "DINKES - SP2TP") = vbYes Then
   strSQL = "DELETE FROM tbPuskesmas WHERE kode='" & Trim(txtData(1).Text) & "'"
   adoCon.Execute strSQL
   adoRec.Requery
   MsgBox "Record Telah Dihapus.", vbInformation, "DINKES - SP2TP"

   If adoRec.BOF And adoRec.EOF Then
    Call Clear
    MsgBox ("Tidak Ada Record Lagi"), vbInformation, "DINKES - SP2TP"
    cmdButton(4).Enabled = False
    cmdButton(1).Enabled = False
    Call Locate
   Else
    adoRec.MoveNext
    If adoRec.EOF Then
     adoRec.MoveLast
    End If
    Call Scatter
   End If
   'message for status of mode
   'lblStatus.Caption = "Record deleted."
  End If
End Sub

Private Sub pSimpan()
On Error GoTo lable

If MsgBox("Apakah data sudah benar ?", vbYesNo + vbQuestion, "DINKES - SP2TP") = vbYes Then
 'Make all entries in input mode enable
 Call SetLock(False)

 'saving new record
 If saveFlag = True Then
  strSQL = "INSERT INTO tbPuskesmas(kode,nama,kdProp,kdKab,kdKec," & _
     "alamat,telp,juml_TT,jml_Pelapor) VALUES('" & Trim(txtData(1).Text) & _
     "','" & Trim(txtData(2).Text) & "','" & Trim(txtData(3).Text) & _
     "','" & Trim(txtData(5).Text) & "','" & Trim(txtData(7).Text) & _
     "','" & Trim(txtData(9).Text) & "','" & Trim(txtData(10).Text) & _
     "'," & Trim(txtData(11).Text) & "," & Trim(txtData(12).Text) & ")"
 Else 'for editing the record
  strSQL = "UPDATE tbPuskesmas SET " & _
     "nama='" & Trim(txtData(2).Text) & _
     "',kdProp='" & Trim(txtData(3).Text) & _
     "',kdKab='" & Trim(txtData(5).Text) & _
     "',kdKec='" & Trim(txtData(7).Text) & _
     "',alamat='" & Trim(txtData(9).Text) & _
     "',telp='" & Trim(txtData(10).Text) & _
     "',juml_TT=" & Trim(txtData(11).Text) & _
     ",jml_Pelapor=" & Trim(txtData(12).Text) & _
     " WHERE kode='" & Trim(txtData(1).Text) & "'"
 End If
 'Make all entries input mode disable
 adoCon.Execute strSQL
 Call SetLock(True)

 adoRec.Requery
 'adoRec.MoveLast
 'show thw current data record
 Call Scatter
 'message for status of mode
 'lblStatus.Caption = " New record Saved."
 MsgBox ("Record Telah Disimpan."), vbInformation + vbOKOnly, "DINKES - SP2TP"
 Call SetButton(True)
 Exit Sub
End If

lable:
 If Err.Number = -2147467259 Then
    MsgBox ("Kode sudah Ada, Isikan dengan Kode Lain"), vbCritical, "DINKES - SP2TP"
    txtData(1).SetFocus
 ElseIf Err.Number <> 0 Then
    MsgBox Err.Number & Err.Description
 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Not exitFlag Then
  Cancel = True
  MsgBox "Gunakan Tombol Tutup", vbOKOnly + vbInformation, "DINKES - SP2TP"
 Else
  Cancel = False
 End If
End Sub

Private Sub sstData_Click(PreviousTab As Integer)
 If cmdButton(5).Enabled Then
  sstData.Tab = 1
 Else
  Call Locate
  Call Scatter
 End If
End Sub

Private Sub txtActivate()
 Dim i As Integer
 
 For i = 1 To 12
  txtData(i).Enabled = True
 Next
 Select Case appTipe
 Case "0"
  For i = 3 To 6
   txtData(i).Enabled = False
  Next
  txtData(3).Text = kdProp
  txtData(4).Text = nmProp
  txtData(5).Text = kdKab
  txtData(6).Text = nmKab
 Case "1"
  For i = 1 To 9
   txtData(i).Enabled = False
  Next
  txtData(1).Text = kdPus
  txtData(2).Text = nmPus
  txtData(3).Text = kdProp
  txtData(4).Text = nmProp
  txtData(5).Text = kdKab
  txtData(6).Text = nmKab
  txtData(7).Text = kdKec
  txtData(8).Text = nmKec
  txtData(9).Text = almt
 Case Else
  For i = 1 To 12
   txtData(i).Enabled = False
  Next
 End Select
End Sub

Private Sub txtdata_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 And Editing Then
  Select Case Index
  Case 3    'Propinsi
   strSQL = "select * from tbProp"
   ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\dinkes07.mdb", strSQL, 1
   txtData(3).Text = Scatter_Code
   txtData(4).Text = Scatter_Code1
  Case 5    'Kabupaten
   strSQL = "select * from tbKab"
   ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\dinkes07.mdb", strSQL, 1
   txtData(5).Text = Scatter_Code
   txtData(6).Text = Scatter_Code1
  Case 7    'Kecamatan
   strSQL = "select * from tbKec"
   ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
            App.Path & "\dinkes07.mdb", strSQL, 1
   txtData(7).Text = Scatter_Code
   txtData(8).Text = Scatter_Code1
  End Select
  SendKeys vbTab
 End If
End Sub
