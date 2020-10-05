VERSION 5.00
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Begin VB.Form rpt103BlThKab 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lap Triwulan 10 Besar Penyakit Kab.Gresik"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
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
      Begin VB.OptionButton optData 
         Caption         =   "TRIWULAN 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "TRIWULAN 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "TRIWULAN 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3480
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optData 
         Caption         =   "TRIWULAN 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Dengan Grafik"
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
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
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
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   5280
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5280
         Y1              =   240
         Y2              =   240
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
         TabIndex        =   2
         Top             =   480
         Width           =   750
      End
   End
End
Attribute VB_Name = "rpt103BlThKab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rsdata As New ADODB.Recordset
Public bl1
Public BlTh As String

Option Explicit

Private Sub cmdButton_Click()
  If optData(0).Value Then
   BlTh = "TRIWULAN PERTAMA"
  ElseIf optData(1).Value Then
   BlTh = "TRIWULAN KEDUA"
  ElseIf optData(2).Value Then
   BlTh = "TRIWULAN KETIGA"
  ElseIf optData(3).Value Then
   BlTh = "TRIWULAN KEEMPAT"
  End If
  
  If chkData.Value = Unchecked Then
   noLap = 7
  Else
   noLap = 8
  End If
  rptForm.Show
End Sub

Private Sub Form_Load()
 pvNum(0).ValueInteger = Year(Date)
End Sub

Private Sub optData_Click(Index As Integer)
 Select Case Index
 Case 0
  bl1 = 1
 Case 1
  bl1 = 4
 Case 2
  bl1 = 7
 Case 3
  bl1 = 10
 End Select
End Sub
