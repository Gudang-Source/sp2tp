VERSION 5.00
Object = "{14BE5479-3D4E-41BE-AF51-F7B42E0FA052}#114.0#0"; "vbComCtl.ocx"
Object = "{C2000000-FFFF-1100-8200-000000000001}#8.0#0"; "PVNum.ocx"
Begin VB.Form frmNTransPenyakit 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Transaksi Penyakit"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   10530
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
      TabIndex        =   47
      Top             =   8760
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
      Left            =   6600
      TabIndex        =   48
      Top             =   8760
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jumlah Usia Terjangkit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   5040
      TabIndex        =   55
      Top             =   2760
      Width           =   5415
      Begin PVNumericLib.PVNumeric pvnum 
         Height          =   300
         Index           =   36
         Left            =   4200
         TabIndex        =   43
         Top             =   4800
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
         Top             =   4440
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
         Top             =   3720
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
         Top             =   3360
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
         Top             =   3000
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
         Top             =   2640
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
         Top             =   2280
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
         Top             =   1560
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
         Top             =   1200
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
         Top             =   840
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
         Top             =   4800
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
         Top             =   1560
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
         Top             =   1560
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
         Top             =   1200
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
         Top             =   1200
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
         Index           =   1
         Left            =   3000
         TabIndex        =   9
         Top             =   840
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
         Top             =   840
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
         Index           =   6
         Left            =   1800
         TabIndex        =   17
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
         Top             =   2280
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
         Top             =   2280
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
         Top             =   2640
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
         Top             =   2640
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
         Top             =   3000
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
         Top             =   3000
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
         Top             =   3360
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
         Top             =   3360
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
         Top             =   3720
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
         Top             =   3720
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
         Top             =   4440
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
         Top             =   4440
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
         Index           =   22
         Left            =   1800
         TabIndex        =   41
         Top             =   4800
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
         Index           =   37
         Left            =   1800
         TabIndex        =   44
         Top             =   5280
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
         Index           =   38
         Left            =   3000
         TabIndex        =   45
         Top             =   5280
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
         TabIndex        =   71
         Top             =   5280
         Width           =   765
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   3960
         Y1              =   5160
         Y2              =   5160
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
         TabIndex        =   70
         Top             =   360
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
         TabIndex        =   69
         Top             =   360
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
         TabIndex        =   68
         Top             =   360
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
         TabIndex        =   67
         Top             =   840
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
         TabIndex        =   66
         Top             =   1200
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
         TabIndex        =   65
         Top             =   1560
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
         TabIndex        =   64
         Top             =   1920
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
         TabIndex        =   63
         Top             =   2280
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
         TabIndex        =   62
         Top             =   2640
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
         TabIndex        =   61
         Top             =   3000
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
         TabIndex        =   60
         Top             =   3360
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
         TabIndex        =   59
         Top             =   3720
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
         TabIndex        =   58
         Top             =   4080
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
         TabIndex        =   57
         Top             =   4440
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
         TabIndex        =   56
         Top             =   4800
         Width           =   1110
      End
   End
   Begin VB.ComboBox cmbData 
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
      Index           =   1
      ItemData        =   "frmNTransPenyakit.frx":0000
      Left            =   120
      List            =   "frmNTransPenyakit.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2760
      Width           =   4815
   End
   Begin VB.ListBox lstData 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5070
      ItemData        =   "frmNTransPenyakit.frx":0004
      Left            =   120
      List            =   "frmNTransPenyakit.frx":0006
      TabIndex        =   7
      Top             =   3360
      Width           =   4815
   End
   Begin vbComCtl.ucFrame ucFrame1 
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4471
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
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   2280
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   2280
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   4560
         TabIndex        =   3
         Top             =   840
         Width           =   5055
      End
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtData 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   5520
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cmbData 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         ItemData        =   "frmNTransPenyakit.frx":0008
         Left            =   2280
         List            =   "frmNTransPenyakit.frx":0030
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2055
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
         Left            =   10080
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Bulan"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   54
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Puskesmas"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Jumlah T.T"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
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
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Jumlah Pelapor"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
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
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   4560
         TabIndex        =   50
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmNTransPenyakit"
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
  For i = 0 To 38
   If i <> 24 Then
    pvNum(i).Enabled = False
    pvNum(i).ValueInteger = 0
   End If
  Next
  lstData.Clear
  If Trim(cmbData(0).Text) <> vbNullString And _
     Trim(txtdata(0).Text) <> vbNullString Then
    Set rsFind = con.Execute("select kode from tbJenisPenyakit " & _
                "where nama='" & Trim(cmbData(1).Text) & "'")
    If Not rsFind.EOF Then
     kdHeader = rsFind(0).Value
    Else
     kdHeader = vbNullString
    End If
    Set rsFind = Nothing
    
    Set rsFind = con.Execute("select kode,nama from tbPenyakit " & _
                 "where kdJenis='" & kdHeader & "'")
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
      lstData.AddItem rsFind(1).Value '& Space(5) & rsFind(0).Value
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
  con.Execute "delete from tbTransPenyakit " & _
     "where kdPuskesmas=''"
  con.Execute "delete from tbTransDtlPenyakit " & _
     "where kdPenyakit=''"
  'Ask = vbNo
  'Call newBtn_Click
  Unload Me
 End Select
End Sub

Private Sub Form_Activate()
 Me.Left = 50
 Me.Top = 50
End Sub

Private Sub Form_Load()
 Set con = New ADODB.Connection
 con.CursorLocation = adUseClient
 'con.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
 '            "Data Source=" & App.Path & "\dinkes07.mdb;"
 con.Open "DSN=dinkesLab"
 cmbData(0).ListIndex = 0
 Set rsFind = con.Execute("select nama from tbJenisPenyakit")
 If Not rsFind.EOF Then rsFind.MoveFirst
 While Not rsFind.EOF
  cmbData(1).AddItem rsFind(0).Value
  rsFind.MoveNext
 Wend
 Set rsFind = Nothing
End Sub

Private Sub lstData_Click()
 Dim i As Integer
 
 If lstData.ListCount > 0 Then
  For i = 0 To 38
   If i <> 24 Then
    pvNum(i).Enabled = True
    pvNum(i).ValueInteger = 0
   End If
  Next
 
  Set rsFind = con.Execute("select * from tbTransDtlPenyakit " & _
               "where no_trans='" & noTrans & "' and " & _
               "kdPenyakit='" & kdKegiatan(lstData.ListIndex) & "'")
  If rsFind.EOF Then
   con.Execute "insert into tbTransDtlPenyakit values('" & _
     noTrans & "','" & kdKegiatan(lstData.ListIndex) & "',0," & _
     "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0," & _
     "0,0,0,0,0,0,0,'',0,0)"
  Else
   pvNum(0).ValueInteger = rsFind("usiaL1").Value
   pvNum(1).ValueInteger = rsFind("usiaP1").Value
   pvNum(25).ValueInteger = rsFind("total1").Value
   pvNum(2).ValueInteger = rsFind("usiaL2").Value
   pvNum(3).ValueInteger = rsFind("usiaP2").Value
   pvNum(26).ValueInteger = rsFind("total2").Value
   pvNum(4).ValueInteger = rsFind("usiaL3").Value
   pvNum(5).ValueInteger = rsFind("usiaP3").Value
   pvNum(27).ValueInteger = rsFind("total3").Value
   pvNum(6).ValueInteger = rsFind("usiaL4").Value
   pvNum(7).ValueInteger = rsFind("usiaP4").Value
   pvNum(28).ValueInteger = rsFind("total4").Value
   pvNum(8).ValueInteger = rsFind("usiaL5").Value
   pvNum(9).ValueInteger = rsFind("usiaP5").Value
   pvNum(29).ValueInteger = rsFind("total5").Value
   pvNum(10).ValueInteger = rsFind("usiaL6").Value
   pvNum(11).ValueInteger = rsFind("usiaP6").Value
   pvNum(30).ValueInteger = rsFind("total6").Value
   pvNum(12).ValueInteger = rsFind("usiaL7").Value
   pvNum(13).ValueInteger = rsFind("usiaP7").Value
   pvNum(31).ValueInteger = rsFind("total7").Value
   pvNum(14).ValueInteger = rsFind("usiaL8").Value
   pvNum(15).ValueInteger = rsFind("usiaP8").Value
   pvNum(32).ValueInteger = rsFind("total8").Value
   pvNum(16).ValueInteger = rsFind("usiaL9").Value
   pvNum(17).ValueInteger = rsFind("usiaP9").Value
   pvNum(33).ValueInteger = rsFind("total9").Value
   pvNum(18).ValueInteger = rsFind("usiaL10").Value
   pvNum(19).ValueInteger = rsFind("usiaP10").Value
   pvNum(34).ValueInteger = rsFind("total10").Value
   pvNum(20).ValueInteger = rsFind("usiaL11").Value
   pvNum(21).ValueInteger = rsFind("usiaP11").Value
   pvNum(35).ValueInteger = rsFind("total11").Value
   pvNum(22).ValueInteger = rsFind("usiaL12").Value
   pvNum(23).ValueInteger = rsFind("usiaP12").Value
   pvNum(36).ValueInteger = rsFind("total12").Value
   pvNum(37).ValueInteger = rsFind("totalL").Value
   pvNum(38).ValueInteger = rsFind("totalP").Value
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
 For i = 0 To 38
  If i <> 24 Then
   pvNum(i).Enabled = False
   pvNum(i).ValueInteger = 0
  End If
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
 For i = 0 To 38
  If i <> 24 Then
   pvNum(i).ValueInteger = 0
   pvNum(i).Enabled = False
  End If
 Next
 lstData.Enabled = False
 lstData.Clear
 cmbData(1).Enabled = False
 txtdata(1).SetFocus
End Sub

Private Sub pvnum_GotFocus(Index As Integer)
 If Index = 24 Then
  On Error Resume Next  'Error Handler berfungsi untuk mengabaikan fokus kontrol
  lstData.SetFocus
 Else
  SendKeys "{home}+{end}"
 End If
End Sub

Private Sub pvnum_KeyPress(Index As Integer, KeyAscii As Integer)
 If KeyAscii = 13 Then
  Select Case Index
  Case 0, 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22 'Laki-Laki
   con.Execute "update tbTransDtlPenyakit set " & _
              "usiaL1=" & pvNum(0).ValueInteger & _
              ", usiaL2=" & pvNum(2).ValueInteger & _
              ", usiaL3=" & pvNum(4).ValueInteger & _
              ", usiaL4=" & pvNum(6).ValueInteger & _
              ", usiaL5=" & pvNum(8).ValueInteger & _
              ", usiaL6=" & pvNum(10).ValueInteger & _
              ", usiaL7=" & pvNum(12).ValueInteger & _
              ", usiaL8=" & pvNum(14).ValueInteger & _
              ", usiaL9=" & pvNum(16).ValueInteger & _
              ", usiaL10=" & pvNum(18).ValueInteger & _
              ", usiaL11=" & pvNum(20).ValueInteger & _
              ", usiaL12=" & pvNum(22).ValueInteger & _
              " where no_trans='" & noTrans & "' AND" & _
              " kdPenyakit='" & kdKegiatan(lstData.ListIndex) & "'"
   If Index = 0 Then pvNum(25).ValueInteger = pvNum(0).ValueInteger + pvNum(1).ValueInteger
   If Index = 2 Then pvNum(26).ValueInteger = pvNum(2).ValueInteger + pvNum(3).ValueInteger
   If Index = 4 Then pvNum(27).ValueInteger = pvNum(4).ValueInteger + pvNum(5).ValueInteger
   If Index = 6 Then pvNum(28).ValueInteger = pvNum(6).ValueInteger + pvNum(7).ValueInteger
   If Index = 8 Then pvNum(29).ValueInteger = pvNum(8).ValueInteger + pvNum(9).ValueInteger
   If Index = 10 Then pvNum(30).ValueInteger = pvNum(10).ValueInteger + pvNum(11).ValueInteger
   If Index = 12 Then pvNum(31).ValueInteger = pvNum(12).ValueInteger + pvNum(13).ValueInteger
   If Index = 14 Then pvNum(32).ValueInteger = pvNum(14).ValueInteger + pvNum(15).ValueInteger
   If Index = 16 Then pvNum(33).ValueInteger = pvNum(16).ValueInteger + pvNum(17).ValueInteger
   If Index = 18 Then pvNum(34).ValueInteger = pvNum(18).ValueInteger + pvNum(19).ValueInteger
   If Index = 20 Then pvNum(35).ValueInteger = pvNum(20).ValueInteger + pvNum(21).ValueInteger
   If Index = 22 Then pvNum(36).ValueInteger = pvNum(22).ValueInteger + pvNum(23).ValueInteger
   
   pvNum(37).ValueInteger = pvNum(0).ValueInteger + pvNum(2).ValueInteger + _
                            pvNum(4).ValueInteger + pvNum(6).ValueInteger + _
                            pvNum(8).ValueInteger + pvNum(10).ValueInteger + _
                            pvNum(12).ValueInteger + pvNum(14).ValueInteger + _
                            pvNum(16).ValueInteger + pvNum(18).ValueInteger + _
                            pvNum(20).ValueInteger + pvNum(22).ValueInteger
  
  Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23    'Perempuan
   con.Execute "update tbTransDtlPenyakit set " & _
              "usiaP1=" & pvNum(1).ValueInteger & _
              ", usiaP2=" & pvNum(3).ValueInteger & _
              ", usiaP3=" & pvNum(5).ValueInteger & _
              ", usiaP4=" & pvNum(7).ValueInteger & _
              ", usiaP5=" & pvNum(9).ValueInteger & _
              ", usiaP6=" & pvNum(11).ValueInteger & _
              ", usiaP7=" & pvNum(13).ValueInteger & _
              ", usiaP8=" & pvNum(15).ValueInteger & _
              ", usiaP9=" & pvNum(17).ValueInteger & _
              ", usiaP10=" & pvNum(19).ValueInteger & _
              ", usiaP11=" & pvNum(21).ValueInteger & _
              ", usiaP12=" & pvNum(23).ValueInteger & _
              " where no_trans='" & noTrans & "' AND" & _
              " kdPenyakit='" & kdKegiatan(lstData.ListIndex) & "'"
   If Index = 1 Then pvNum(25).ValueInteger = pvNum(0).ValueInteger + pvNum(1).ValueInteger
   If Index = 3 Then pvNum(26).ValueInteger = pvNum(2).ValueInteger + pvNum(3).ValueInteger
   If Index = 5 Then pvNum(27).ValueInteger = pvNum(4).ValueInteger + pvNum(5).ValueInteger
   If Index = 7 Then pvNum(28).ValueInteger = pvNum(6).ValueInteger + pvNum(7).ValueInteger
   If Index = 9 Then pvNum(29).ValueInteger = pvNum(8).ValueInteger + pvNum(9).ValueInteger
   If Index = 11 Then pvNum(30).ValueInteger = pvNum(10).ValueInteger + pvNum(11).ValueInteger
   If Index = 13 Then pvNum(31).ValueInteger = pvNum(12).ValueInteger + pvNum(13).ValueInteger
   If Index = 15 Then pvNum(32).ValueInteger = pvNum(14).ValueInteger + pvNum(15).ValueInteger
   If Index = 17 Then pvNum(33).ValueInteger = pvNum(16).ValueInteger + pvNum(17).ValueInteger
   If Index = 19 Then pvNum(34).ValueInteger = pvNum(18).ValueInteger + pvNum(19).ValueInteger
   If Index = 21 Then pvNum(35).ValueInteger = pvNum(20).ValueInteger + pvNum(21).ValueInteger
   If Index = 23 Then pvNum(36).ValueInteger = pvNum(22).ValueInteger + pvNum(23).ValueInteger
   
   pvNum(38).ValueInteger = pvNum(1).ValueInteger + pvNum(3).ValueInteger + _
                            pvNum(5).ValueInteger + pvNum(7).ValueInteger + _
                            pvNum(9).ValueInteger + pvNum(11).ValueInteger + _
                            pvNum(13).ValueInteger + pvNum(15).ValueInteger + _
                            pvNum(17).ValueInteger + pvNum(19).ValueInteger + _
                            pvNum(21).ValueInteger + pvNum(23).ValueInteger
  
  Case 25 To 36   'Total1 to Total12
   con.Execute "update tbTransDtlPenyakit set " & _
              "total1=" & pvNum(25).ValueInteger & _
              ", total2=" & pvNum(26).ValueInteger & _
              ", total3=" & pvNum(27).ValueInteger & _
              ", total4=" & pvNum(28).ValueInteger & _
              ", total5=" & pvNum(29).ValueInteger & _
              ", total6=" & pvNum(30).ValueInteger & _
              ", total7=" & pvNum(31).ValueInteger & _
              ", total8=" & pvNum(32).ValueInteger & _
              ", total9=" & pvNum(33).ValueInteger & _
              ", total10=" & pvNum(34).ValueInteger & _
              ", total11=" & pvNum(35).ValueInteger & _
              ", total12=" & pvNum(36).ValueInteger & _
              " where no_trans='" & noTrans & "' AND" & _
              " kdPenyakit='" & kdKegiatan(lstData.ListIndex) & "'"
  
  Case 37, 38   'TotalL & TotalP
   con.Execute "update tbTransDtlPenyakit set " & _
              "totalL=" & pvNum(37).ValueInteger & _
              ", totalP=" & pvNum(38).ValueInteger & _
              " where no_trans='" & noTrans & "' AND" & _
              " kdPenyakit='" & kdKegiatan(lstData.ListIndex) & "'"
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
    con.Execute "update tbTransPenyakit set " & _
        "jumlahTT=" & val(Trim(txtdata(3).Text)) & _
        " where no_trans='" & noTrans & "'"
   Else
    MsgBox "Pastikan Kolom Jumlah T.T sudah terisi", vbOKOnly + vbInformation, "DINAS KESEHATAN"
    txtdata(1).SetFocus 'Set Fokus ke Kolom Jumlah T.T
   End If
  
  Case 4
   If Trim(txtdata(4).Text) <> vbNullString Then
    con.Execute "update tbTransPenyakit set " & _
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
 
 Ask = MsgBox("Data LB1 Baru ?", vbYesNoCancel + vbQuestion + vbDefaultButton1, "DINAS KESEHATAN")
 If Ask = vbYes Then        'Data Baru
   Set rsFind = con.Execute("select * from tbTransPenyakit " & _
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
   
   con.Execute "insert into tbTransPenyakit values('" & noTrans & _
       "'," & cmbData(0).ListIndex + 1 & "," & val(txtdata(0).Text) & _
       ",'" & Trim(txtdata(1).Text) & "',0,0)"
   
   txtdata(3).Enabled = True
   txtdata(4).Enabled = True
   
   cmbData(1).Enabled = True
   lstData.Enabled = True
 ElseIf Ask = vbNo Then     'Data Lama
  MySql = "select kdPuskesmas,namaPus,jumlahTT,pelapor,no_trans from qDistincTransPus " & _
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
    For i = 0 To 38
     If i <> 24 Then pvNum(i).Enabled = True
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
  pvNum(24).SetFocus 'Set Fokus ke Combo Bulan
 End If
End Sub
