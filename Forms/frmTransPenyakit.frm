VERSION 5.00
Object = "{14BE5479-3D4E-41BE-AF51-F7B42E0FA052}#114.0#0"; "vbComCtl.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Begin VB.Form frmTransPenyakit 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Transaksi Penyakit"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   8745
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
      TabIndex        =   71
      Top             =   7920
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
      Left            =   6240
      TabIndex        =   70
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jumlah Usia Terjangkit"
      Height          =   5655
      Left            =   3240
      TabIndex        =   53
      Top             =   2160
      Width           =   5415
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   38
         Left            =   3000
         TabIndex        =   45
         Top             =   5160
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
         Index           =   37
         Left            =   1800
         TabIndex        =   44
         Top             =   5160
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
         Index           =   24
         Left            =   4560
         TabIndex        =   46
         Top             =   6480
         Width           =   585
         _Version        =   524288
         _ExtentX        =   1032
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
         Index           =   22
         Left            =   1800
         TabIndex        =   41
         Top             =   4680
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
         Index           =   21
         Left            =   3000
         TabIndex        =   39
         Top             =   4320
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
         Index           =   20
         Left            =   1800
         TabIndex        =   38
         Top             =   4320
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
         Index           =   19
         Left            =   3000
         TabIndex        =   36
         Top             =   3960
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
         Index           =   18
         Left            =   1800
         TabIndex        =   35
         Top             =   3960
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
         Index           =   17
         Left            =   3000
         TabIndex        =   33
         Top             =   3600
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
         Index           =   16
         Left            =   1800
         TabIndex        =   32
         Top             =   3600
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
         Index           =   15
         Left            =   3000
         TabIndex        =   30
         Top             =   3240
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
         Index           =   14
         Left            =   1800
         TabIndex        =   29
         Top             =   3240
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
         Index           =   13
         Left            =   3000
         TabIndex        =   27
         Top             =   2880
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
         Index           =   12
         Left            =   1800
         TabIndex        =   26
         Top             =   2880
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
         Index           =   11
         Left            =   3000
         TabIndex        =   24
         Top             =   2520
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
         Index           =   10
         Left            =   1800
         TabIndex        =   23
         Top             =   2520
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
         Index           =   9
         Left            =   3000
         TabIndex        =   21
         Top             =   2160
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
         Index           =   8
         Left            =   1800
         TabIndex        =   20
         Top             =   2160
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
         Index           =   7
         Left            =   3000
         TabIndex        =   18
         Top             =   1800
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
         Index           =   6
         Left            =   1800
         TabIndex        =   17
         Top             =   1800
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
         Left            =   1800
         TabIndex        =   8
         Top             =   720
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
         LimitValue      =   -1  'True
         DisableSpins    =   -1  'True
      End
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   1
         Left            =   3000
         TabIndex        =   9
         Top             =   720
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
         Left            =   1800
         TabIndex        =   11
         Top             =   1080
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
         Index           =   3
         Left            =   3000
         TabIndex        =   12
         Top             =   1080
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
         Index           =   4
         Left            =   1800
         TabIndex        =   14
         Top             =   1440
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
         Index           =   5
         Left            =   3000
         TabIndex        =   15
         Top             =   1440
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
         Index           =   23
         Left            =   3000
         TabIndex        =   42
         Top             =   4680
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
         Index           =   25
         Left            =   4200
         TabIndex        =   10
         Top             =   720
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
         Index           =   26
         Left            =   4200
         TabIndex        =   13
         Top             =   1080
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
         Index           =   27
         Left            =   4200
         TabIndex        =   16
         Top             =   1440
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
         Index           =   28
         Left            =   4200
         TabIndex        =   19
         Top             =   1800
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
         Index           =   29
         Left            =   4200
         TabIndex        =   22
         Top             =   2160
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
         Index           =   30
         Left            =   4200
         TabIndex        =   25
         Top             =   2520
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
         Index           =   31
         Left            =   4200
         TabIndex        =   28
         Top             =   2880
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
         Index           =   32
         Left            =   4200
         TabIndex        =   31
         Top             =   3240
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
         Index           =   33
         Left            =   4200
         TabIndex        =   34
         Top             =   3600
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
         Index           =   34
         Left            =   4200
         TabIndex        =   37
         Top             =   3960
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
         Index           =   35
         Left            =   4200
         TabIndex        =   40
         Top             =   4320
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
         Index           =   36
         Left            =   4200
         TabIndex        =   43
         Top             =   4680
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
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
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
         Left            =   240
         TabIndex        =   69
         Top             =   5160
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   3960
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL"
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
         Left            =   4200
         TabIndex        =   68
         Top             =   240
         Width           =   765
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
         Left            =   2880
         TabIndex        =   67
         Top             =   240
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
         Left            =   1800
         TabIndex        =   66
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "0 - 7 Hari"
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
         TabIndex        =   65
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "8 - 28 Hari"
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
         TabIndex        =   64
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "1 Bulan < 1 th"
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
         TabIndex        =   63
         Top             =   1440
         Width           =   1320
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "1 - 4 Tahun"
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
         TabIndex        =   62
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "5 - 9 tahun"
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
         TabIndex        =   61
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "10 - 14 Tahun"
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
         TabIndex        =   60
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "15 - 19 Tahun"
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
         TabIndex        =   59
         Top             =   2880
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "20 - 44 Tahun"
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
         TabIndex        =   58
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "45 - 54 Tahun"
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
         TabIndex        =   57
         Top             =   3600
         Width           =   1305
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "55 - 59 tahun"
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
         TabIndex        =   56
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "60 - 69 tahun"
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
         TabIndex        =   55
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   ">= 70 tahun"
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
         TabIndex        =   54
         Top             =   4680
         Width           =   1110
      End
   End
   Begin VB.ComboBox cmbData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      ItemData        =   "frmTransPenyakit.frx":0000
      Left            =   120
      List            =   "frmTransPenyakit.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ListBox lstData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   5295
      ItemData        =   "frmTransPenyakit.frx":0004
      Left            =   120
      List            =   "frmTransPenyakit.frx":0006
      TabIndex        =   7
      Top             =   2520
      Width           =   3015
   End
   Begin vbComCtl.ucFrame ucFrame1 
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
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
      Caption         =   "Entry Data Transaksi"
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1920
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
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   3120
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   4680
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbData 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmTransPenyakit.frx":0008
         Left            =   1920
         List            =   "frmTransPenyakit.frx":0030
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
         Left            =   8280
         TabIndex        =   47
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   1440
         Width           =   1815
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
         Left            =   3720
         TabIndex        =   48
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmTransPenyakit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim rsFind As New ADODB.Recordset
Dim rsFind2 As New ADODB.Recordset
Dim strCon As String, tEdit As Boolean
Public noTrans As String
Public swShowTr As Boolean
Dim kdPenyakit() As String
Public Ask As Variant

Option Explicit

Private Sub cmbData_Click(Index As Integer)
 Dim i As Integer, kdJenis As String
 
 Select Case Index
 Case 1 'Pilih Jenis Penyakit
  lstData.Clear
  If Trim(cmbData(0).Text) <> vbNullString And _
     Trim(txtData(0).Text) <> vbNullString Then
   'Select Case cmbData(1).ListIndex
   'Case 0
    i = 0
    Set rsFind = con.Execute("select kode from tbJenisPenyakit " & _
                "where nama='" & Trim(cmbData(1).Text) & "'")
    If Not rsFind.EOF Then
     kdJenis = rsFind(0).Value
    Else
     kdJenis = vbNullString
    End If
    Set rsFind = Nothing
    
    Set rsFind = con.Execute("select kode,nama from tbPenyakit " & _
                 "where kdJenis='" & kdJenis & "'")
    If Not rsFind.EOF Then
     rsFind.MoveFirst
     While Not rsFind.EOF
      rsFind.MoveNext
      i = i + 1
     Wend
     rsFind.MoveFirst
     ReDim kdPenyakit(i) As String
     i = 0
     While Not rsFind.EOF
      lstData.AddItem rsFind(0).Value & Space(5) & rsFind(1).Value
      kdPenyakit(i) = rsFind(0).Value
      rsFind.MoveNext
      i = i + 1
     Wend
    End If
    Set rsFind = Nothing
   'End Select
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
   
   Set rsFind = con.Execute("select * from tbTransPenyakit " & _
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
   
   If appTipe = "0" Or appTipe = "2" Then
    txtData(1).Enabled = True
    txtData(3).Enabled = True
    txtData(4).Enabled = True
   Else
    txtData(1).Enabled = False
    txtData(3).Enabled = False
    txtData(4).Enabled = False
   End If
   cmbData(1).Enabled = True
   lstData.Enabled = True
   
   Ask = MsgBox("Data LB1 Baru ?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "DINKES")
   If Ask = vbYes Then
    tEdit = False
    clearForm
    con.BeginTrans
    If appTipe <> "1" Then
     con.Execute "insert into tbTransPenyakit values('" & noTrans & _
      "'," & cmbData(0).ListIndex + 1 & "," & val(txtData(0).Text) & ",'',0,'')"
    Else
     con.Execute "insert into tbTransPenyakit values('" & noTrans & _
      "'," & cmbData(0).ListIndex + 1 & "," & _
      val(txtData(0).Text) & ",'" & kdPus & _
      "'," & txtData(3).Text & ",'" & _
      txtData(4).Text & "')"
    End If
    con.CommitTrans
    If appTipe <> "1" Then txtData(1).SetFocus
   ElseIf Ask = vbNo Then
    tEdit = True
    noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
    noTrans = noTrans & Right(txtData(0).Text, 2) & "-"
    
    MySql = "select kdPuskesmas,namaPus,jumlahTT,pelapor,no_trans from QDistincTransPus " & _
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
  
 Case 1  'List Jenis Penyakit
   If Trim(cmbData(0).Text) <> vbNullString And _
      Trim(txtData(0).Text) <> vbNullString And _
      Ask = vbYes Then
      
    On Error Resume Next
    con.BeginTrans
    con.Execute "insert into tbTransDtlPenyakit values('" & _
       noTrans & "','',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & _
       "0,0,0,0,0,0,'',0,0)"
    con.CommitTrans
   End If
 End Select
End Sub

Private Sub cmdBtn_Click(Index As Integer)
 Select Case Index
 Case 0
  con.Execute "delete from tbTransPenyakit " & _
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
 
 Set rsFind = con.Execute("select nama from tbJenisPenyakit")
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
 
 For i = 0 To 23
  pvnum(i).ValueInteger = 0
 Next
 For i = 25 To 38
  pvnum(i).ValueInteger = 0
 Next
 nextBtn.Enabled = False
 newBtn.Enabled = False
End Sub

Private Sub lstData_Click()
 Dim i As Integer
 
 If Trim(cmbData(0).Text) <> vbNullString And _
    Trim(txtData(0).Text) <> vbNullString And _
    lstData.ListCount > 0 Then
 
  If Ask = vbYes Then
   con.BeginTrans
   con.Execute "update tbTransDtlPenyakit set " & _
               "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & _
               "' where no_Trans='" & noTrans & "'"
   con.CommitTrans
   Ask = vbNo
   If appTipe <> "1" Then nextBtn.Enabled = True
  ElseIf Ask = vbNo Then
   Set rsFind = con.Execute("select * from tbTransDtlPenyakit " & _
                "where no_trans='" & noTrans & "' and " & _
                "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'")
   If rsFind.EOF Then
    con.BeginTrans
    con.Execute "insert into tbTransDtlPenyakit values('" & _
       noTrans & "','" & kdPenyakit(lstData.ListIndex) & "',0," & _
       "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'',0,0)"
    con.CommitTrans
   End If
   Set rsFind = Nothing
  End If
 
  For i = 0 To 23
   pvnum(i).ValueInteger = 0
   pvnum(i).Enabled = True
  Next
  For i = 25 To 38
   pvnum(i).ValueInteger = 0
   pvnum(i).Enabled = True
  Next
  
 
     'Set rsFind = con.Execute("select usia1,usia2,usia3,usia4,usia5,usia6," & _
               "usia7,usia8,usia9,usia10,usia11,usia12 from tbPenyakit " & _
               "where kode='" & kdPenyakit(lstData.ListIndex) & "'")
    'If Not rsFind.EOF Then
   ' For i = 0 To 11
   '  If Not rsFind(i).Value Then
   '   pvnum(i * 2).Enabled = False
   '   pvnum((i * 2) + 1).Enabled = False
   '   pvnum(i + 25).Enabled = False
   '  End If
   ' Next
   'End If
   'Set rsFind = Nothing
  
  Set rsFind = con.Execute("select usiaL1,usiaP1,usiaL2,usiaP2," & _
               "usiaL3,usiaP3,usiaL4,usiaP4,usiaL5,usiaP5,usiaL6," & _
               "usiaP6,usiaL7,usiaP7,usiaL8,usiaP8,usiaL9,usiaP9," & _
               "usiaL10,usiaP10,usiaL11,usiaP11,usiaL12,usiaP12," & _
               "total1,total2,total3,total4,total5,total6,total7," & _
               "total8,total9,total10,total11,total12,totalL,totalP from qTransPenyakit " & _
               "where kdPenyakit='" & kdPenyakit(lstData.ListIndex) & _
               "' and no_trans='" & noTrans & "'")
  If Not rsFind.EOF Then
   For i = 0 To 23
    pvnum(i).ValueInteger = rsFind(i).Value
   Next
   For i = 24 To 37
    pvnum(i + 1).ValueInteger = rsFind(i).Value
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
 For i = 1 To 4
  txtData(i).Enabled = False
 Next
 cmbData(1).Enabled = False
 lstData.Enabled = False
 
 txtData(0).Text = vbNullString
 cmbData(0).Enabled = True
 txtData(0).Enabled = True
 cmbData(0).SetFocus
' Ask = ""
 noTrans = vbNullString
 newBtn.Enabled = False
 nextBtn.Enabled = False
End Sub

Private Sub nextBtn_Click()
 Dim no As Variant
 Dim i As Integer, nul As String
 
 clearForm
' If Ask = vbYes Then
 If Not tEdit Then
  noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
  noTrans = noTrans & Right(txtData(0).Text, 2) & "-"
   
  Set rsFind = con.Execute("select * from tbTransPenyakit " & _
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
  con.BeginTrans
  con.Execute "insert into tbTransPenyakit values('" & noTrans & _
   "'," & cmbData(0).ListIndex + 1 & "," & val(txtData(0).Text) & ",'',0,'')"
  con.CommitTrans
 End If
 txtData(1).SetFocus
 'nextBtn.Enabled = False
End Sub

Private Sub pvnum_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub totalLP()
 pvnum(37).Text = pvnum(0).ValueInteger + pvnum(2).ValueInteger + pvnum(4).ValueInteger + _
         pvnum(6).ValueInteger + pvnum(8).ValueInteger + pvnum(10).ValueInteger + _
         pvnum(12).ValueInteger + pvnum(14).ValueInteger + pvnum(16).ValueInteger + _
         pvnum(18).ValueInteger + pvnum(20).ValueInteger + pvnum(22).ValueInteger
 pvnum(38).Text = pvnum(1).ValueInteger + pvnum(3).ValueInteger + pvnum(5).ValueInteger + _
         pvnum(7).ValueInteger + pvnum(9).ValueInteger + pvnum(11).ValueInteger + _
         pvnum(13).ValueInteger + pvnum(15).ValueInteger + pvnum(17).ValueInteger + _
         pvnum(19).ValueInteger + pvnum(21).ValueInteger + pvnum(23).ValueInteger
End Sub

Private Sub pvnum_lostFocus(Index As Integer)
 Select Case Index
 Case 0, 1, 25
  If Index <> 25 Then
   totalLP
   pvnum(25).Text = pvnum(0).ValueInteger + pvnum(1).ValueInteger
  End If
  On Error Resume Next
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL1=" & pvnum(0).ValueInteger & ",usiaP1=" & pvnum(1).ValueInteger & _
              ",total1=" & pvnum(25).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 2, 3, 26
  If Index <> 26 Then
   totalLP
   pvnum(26).Text = pvnum(2).ValueInteger + pvnum(3).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL2=" & pvnum(2).ValueInteger & ",usiaP2=" & pvnum(3).ValueInteger & _
              ",total2=" & pvnum(26).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 4, 5, 27
  If Index <> 27 Then
   totalLP
   pvnum(27).Text = pvnum(4).ValueInteger + pvnum(5).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL3=" & pvnum(4).ValueInteger & ",usiaP3=" & pvnum(5).ValueInteger & _
              ",total3=" & pvnum(27).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 6, 7, 28
  If Index <> 28 Then
   totalLP
   pvnum(28).Text = pvnum(6).ValueInteger + pvnum(7).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL4=" & pvnum(6).ValueInteger & ",usiaP4=" & pvnum(7).ValueInteger & _
              ",total4=" & pvnum(28).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 8, 9, 29
  If Index <> 29 Then
   totalLP
   pvnum(29).Text = pvnum(8).ValueInteger + pvnum(9).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL5=" & pvnum(8).ValueInteger & ",usiaP5=" & pvnum(9).ValueInteger & _
              ",total5=" & pvnum(29).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 10, 11, 30
  If Index <> 30 Then
   totalLP
   pvnum(30).Text = pvnum(10).ValueInteger + pvnum(11).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL6=" & pvnum(10).ValueInteger & ",usiaP6=" & pvnum(11).ValueInteger & _
              ",total6=" & pvnum(30).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 12, 13, 31
  If Index <> 31 Then
   totalLP
   pvnum(31).Text = pvnum(12).ValueInteger + pvnum(13).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL7=" & pvnum(12).ValueInteger & ",usiaP7=" & pvnum(13).ValueInteger & _
              ",total7=" & pvnum(31).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 14, 15, 32
  If Index <> 32 Then
   totalLP
   pvnum(32).Text = pvnum(14).ValueInteger + pvnum(15).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL8=" & pvnum(14).ValueInteger & ",usiaP8=" & pvnum(15).ValueInteger & _
              ",total8=" & pvnum(32).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 16, 17, 33
  If Index <> 33 Then
   totalLP
   pvnum(33).Text = pvnum(16).ValueInteger + pvnum(17).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL9=" & pvnum(16).ValueInteger & ",usiaP9=" & pvnum(17).ValueInteger & _
              ",total9=" & pvnum(33).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 18, 19, 34
  If Index <> 34 Then
   totalLP
   pvnum(34).Text = pvnum(18).ValueInteger + pvnum(19).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL10=" & pvnum(18).ValueInteger & ",usiaP10=" & pvnum(19).ValueInteger & _
              ",total10=" & pvnum(34).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 20, 21, 35
  If Index <> 35 Then
   totalLP
   pvnum(35).Text = pvnum(20).ValueInteger + pvnum(21).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL11=" & pvnum(20).ValueInteger & ",usiaP11=" & pvnum(21).ValueInteger & _
              ",total11=" & pvnum(35).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 22, 23, 36
  If Index <> 36 Then
   totalLP
   pvnum(36).Text = pvnum(22).ValueInteger + pvnum(23).ValueInteger
  End If
  con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL12=" & pvnum(22).ValueInteger & ",usiaP12=" & pvnum(23).ValueInteger & _
              ",total12=" & pvnum(36).ValueInteger & " where no_trans='" & noTrans & "' AND " & _
              "kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 Case 37, 38
  con.Execute "update tbTransDtlPenyakit set " & _
              "totalL=" & pvnum(37).ValueInteger & _
              ",totalP=" & pvnum(38).ValueInteger & _
              " where no_trans='" & noTrans & "' AND " & _
              " kdPenyakit='" & kdPenyakit(lstData.ListIndex) & "'"
 End Select
 If Index = 38 Then lstData.SetFocus
End Sub

Private Sub txtData_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim MySql As String
  
  If KeyAscii = 13 Then
   Select Case Index
   Case 0
    Call cmbData_LostFocus(0)
    If noTrans <> vbNullString Then
     cmbData(0).Enabled = False
     txtData(0).Enabled = False
     newBtn.Enabled = True
    Else
     txtData(1).Enabled = False
     txtData(3).Enabled = False
     txtData(4).Enabled = False
     cmbData(1).Enabled = False
     lstData.Enabled = False
    End If
   Case 1
    If Trim(txtData(1).Text) = vbNullString Then
     'If Ask = vbYes Then
     If Not tEdit Then
      If appTipe <> "1" Then
       MySql = "select kode,nama from tbPuskesmas"
       ShowFind "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
               App.Path & "\dinkes07.mdb", MySql, 1
       txtData(1).Text = Scatter_Code
       txtData(2).Text = Scatter_Code1
      End If
      SendKeys vbTab
     'ElseIf Ask = vbNo Then
     ElseIf tEdit Then
      noTrans = IIf(Len(Trim(Str(cmbData(0).ListIndex + 1))) <> 2, "0" & Trim(Str(cmbData(0).ListIndex + 1)), Trim(Str(cmbData(0).ListIndex + 1)))
      noTrans = noTrans & Right(txtData(0).Text, 2) & "-"
   
      MySql = "select kdPuskesmas,namaPus,jumlahTT,pelapor,no_trans from qDistincTransPus " & _
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
   con.Execute "update tbTransPenyakit set " & _
               "kdPuskesmas='" & txtData(1).Text & _
               "' where no_Trans='" & noTrans & "'"
  End If
 Case 3
  con.Execute "update tbTransPenyakit set " & _
              "jumlahTT=" & val(txtData(3).Text) & _
              " where no_Trans='" & noTrans & "'"
 Case 4
  con.Execute "update tbTransPenyakit set " & _
              "pelapor='" & txtData(4).Text & _
              "' where no_Trans='" & noTrans & "'"
 End Select
End Sub

Private Sub clearForm()
 Dim i As Integer
 
 For i = 0 To 23
  pvnum(i).ValueInteger = 0
  pvnum(i).Enabled = False
 Next
 For i = 25 To 38
  pvnum(i).ValueInteger = 0
  pvnum(i).Enabled = False
 Next
 
 If appTipe <> "1" Then
  For i = 1 To 4
   txtData(i).Text = vbNullString
  Next
 End If
End Sub
