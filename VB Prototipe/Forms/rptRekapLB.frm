VERSION 5.00
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "pvnum.ocx"
Begin VB.Form rptRekapLB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Rekap LB per Tahun"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin PVNumericLib.PVNumeric pvNum 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   1095
         _Version        =   524288
         _ExtentX        =   1931
         _ExtentY        =   529
         _StockProps     =   253
         Text            =   "0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         Alignment       =   2
         LimitValue      =   -1  'True
         DecimalMax      =   0
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   1
         ItemData        =   "rptRekapLB.frx":0000
         Left            =   1560
         List            =   "rptRekapLB.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdButton 
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   5
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cmbData 
         Height          =   315
         Index           =   0
         ItemData        =   "rptRekapLB.frx":0004
         Left            =   3480
         List            =   "rptRekapLB.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
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
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1125
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   5280
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5280
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LB"
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
         Height          =   240
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tahun :"
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
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   750
      End
   End
End
Attribute VB_Name = "rptRekapLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rsdata As New ADODB.Recordset
Public cKode As String

Option Explicit

Private Sub cmbData_Click(Index As Integer)
 Select Case Index
 Case 0
   cmbData(1).Clear
   Select Case cmbData(0).ListIndex
   Case 0
    noLap = 10
    Set rsdata = con.Execute("select * from tbPenyakit order by kode")
   Case 1
    noLap = 11
    Set rsdata = con.Execute("select * from tbGKIA order by kode")
   Case 2
    noLap = 12
    Set rsdata = con.Execute("select * from tbSubKegiatan order by kode")
   End Select
   If Not rsdata.EOF Then rsdata.MoveFirst
   While Not rsdata.EOF
    cmbData(1).AddItem rsdata("kode").Value & " ( " & rsdata("nama").Value & " )"
    rsdata.MoveNext
   Wend
   Set rsdata = Nothing
   
 Case 1
   cKode = Left(Trim(cmbData(1).Text), 10)
 End Select
End Sub

Private Sub cmdButton_Click()
  rptForm.Show
End Sub

Private Sub Form_Load()
 pvNum(0).ValueInteger = Year(Date)
  
 con.CursorLocation = adUseClient
 con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & App.Path & "\dinkes07.mdb;"
             
End Sub

Private Sub Form_Unload(Cancel As Integer)
 con.Close
End Sub
