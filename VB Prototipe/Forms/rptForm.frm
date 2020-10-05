VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rptForm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "rptForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report0 As New rptPenyakit
Dim Report1 As New rptKunjungan
Dim Report2 As New rptGKIA
Dim Report3 As New rptPuskesmas
Dim Report4 As New rpt10Besar
Dim Report5 As New rptChart10Besar
Dim Report6 As New rptTw1Chart10Besar   'Triwulan 1
Dim Report7 As New rptTw2Chart10Besar   'Triwulan 2
Dim Report8 As New rptTw3Chart10Besar   'Triwulan 3
Dim Report9 As New rptTw4Chart10Besar   'Triwulan 3
Dim Report10 As New rpt10BesarPerTahun
Dim Report11 As New rptRekapLB1
Dim Report12 As New rpt10BesarP
Dim Report13 As New rpt10BesarPerTahunP

Private Sub Form_Load()
 Dim BlTh As String
 
 Screen.MousePointer = vbHourglass
 Select Case noLap
 Case 1
  'If rptLB1.chkdata.Value = Checked Then
  ' CRViewer1.ReportSource = Report4
  ' Report4.PrinterSetup Me.hWnd
  'Else
   Report0.FormulaFields(1).Text = "'" & rptLB1.cmbData(0).Text & "'"
   Report0.FormulaFields(2).Text = "'" & rptLB1.pvNum(0).Text & "'"
  
   Report0.RecordSelectionFormula = "{qTransPenyakit.bulan}= " & _
      rptLB1.cmbData(0).ListIndex + 1 & " and {qTransPenyakit.tahun}=" & _
      rptLB1.pvNum(0).ValueInteger & ""
     
   If rptLB1.cmbData(1).ListIndex <> 0 Then
    Report0.RecordSelectionFormula = Report0.RecordSelectionFormula & _
      " and {qTransPenyakit.kdKec}='" & Trim(Left(rptLB1.cmbData(1).Text, 3)) & "'"
   End If
  
   If rptLB1.cmbData(2).ListIndex <> 0 Then
    Report0.RecordSelectionFormula = Report0.RecordSelectionFormula & _
      " and {qTransPenyakit.kdPuskesmas}='" & Trim(rptLB1.cmbData(2).Text) & "'"
   End If
  
   CRViewer1.ReportSource = Report0
   Report0.PrinterSetup Me.hWnd
  'End If
    
 Case 2
  Report2.FormulaFields(1).Text = "'" & rptLB3.cmbData(0).Text & "'"
  Report2.FormulaFields(2).Text = "'" & rptLB3.pvNum(0).Text & "'"
  
  Report2.RecordSelectionFormula = "{qTransGKIA.bulan}= " & _
     rptLB3.cmbData(0).ListIndex + 1 & " and {qTransGKIA.tahun}=" & _
     rptLB3.pvNum(0).ValueInteger & ""
     
  If rptLB3.cmbData(1).ListIndex <> 0 Then
   Report2.RecordSelectionFormula = Report2.RecordSelectionFormula & _
     " and {qTransGKIA.kdKec}='" & Trim(Left(rptLB3.cmbData(1).Text, 3)) & "'"
  End If
  
  If rptLB3.cmbData(2).ListIndex <> 0 Then
   Report2.RecordSelectionFormula = Report2.RecordSelectionFormula & _
     " and {qTransGKIA.kdPuskesmas}='" & Trim(rptLB3.cmbData(2).Text) & "'"
  End If
  
  CRViewer1.ReportSource = Report2
  Report2.PrinterSetup Me.hWnd
  
 Case 3
  Report1.FormulaFields(1).Text = "'" & rptLB4.cmbData(0).Text & "'"
  Report1.FormulaFields(2).Text = "'" & rptLB4.pvNum(0).Text & "'"
  
  Report1.RecordSelectionFormula = "{qTransKegiatan.bulan}= " & _
     rptLB4.cmbData(0).ListIndex + 1 & " and {qTransKegiatan.tahun}=" & _
     rptLB4.pvNum(0).ValueInteger & ""
     
  If rptLB4.cmbData(1).ListIndex <> 0 Then
   Report1.RecordSelectionFormula = Report1.RecordSelectionFormula & _
     " and {qTransKegiatan.kdKec}='" & Trim(Left(rptLB4.cmbData(1).Text, 3)) & "'"
  End If
  
  If rptLB4.cmbData(2).ListIndex <> 0 Then
   Report1.RecordSelectionFormula = Report1.RecordSelectionFormula & _
     " and {qTransKegiatan.kdPuskesmas}='" & Trim(rptLB4.cmbData(2).Text) & "'"
  End If
   
  CRViewer1.ReportSource = Report1
  Report1.PrinterSetup Me.hWnd
  
 Case 4
  Report3.PrinterSetup Me.hWnd
  CRViewer1.ReportSource = Report3
  
 Case 5
  BlTh = UCase(rpt10BlThKab.cmbData(0).Text) & _
         " " & Trim(Str(rpt10BlThKab.pvNum(0).ValueReal))
  Report4.FormulaFields(18).Text = "'" & BlTh & "'"
  
  Report4.RecordSelectionFormula = "{qryTopPenyakit.bulan}= " & _
     rpt10BlThKab.cmbData(0).ListIndex + 1 & " and {qryTopPenyakit.tahun}=" & _
     rpt10BlThKab.pvNum(0).ValueInteger & ""
  
  Report4.PrinterSetup Me.hWnd
  CRViewer1.ReportSource = Report4
  
 Case 6
  BlTh = UCase(rpt10BlThKab.cmbData(0).Text) & _
         " " & Trim(Str(rpt10BlThKab.pvNum(0).ValueReal))
  Report5.FormulaFields(18).Text = "'" & BlTh & "'"
  
  Report5.RecordSelectionFormula = "{qryTopPenyakit.bulan}= " & _
     rpt10BlThKab.cmbData(0).ListIndex + 1 & " and {qryTopPenyakit.tahun}=" & _
     rpt10BlThKab.pvNum(0).ValueInteger & ""
     
  Report5.PrinterSetup Me.hWnd
  CRViewer1.ReportSource = Report5
 
 'Case 7
  'Report4.FormulaFields(18).Text = "'" & rpt103BlThKab.BlTh & "'"
    
  'Report4.RecordSelectionFormula = "({qryTopPenyakit.bulan}>= " & _
    rpt103BlThKab.bl1 & " and {qryTopPenyakit.bulan}<= " & _
    rpt103BlThKab.bl2 & ") and {qryTopPenyakit.tahun}=" & _
    rpt103BlThKab.pvNum(0).ValueInteger & ""
    
  'Report4.PrinterSetup Me.hWnd
  'CRViewer1.ReportSource = Report4
  
 Case 7, 8 'Laporan 10 Besar Per Triwulan
  Select Case rpt103BlThKab.bl1
  Case 1
   Report6.FormulaFields(18).Text = "'" & rpt103BlThKab.BlTh & "'"
    
   Report6.RecordSelectionFormula = "{qryTopTw1Penyakit.tahun}=" & _
    rpt103BlThKab.pvNum(0).ValueInteger & ""
  
   Report6.PrinterSetup Me.hWnd
   CRViewer1.ReportSource = Report6
  
  Case 4
   Report7.FormulaFields(18).Text = "'" & rpt103BlThKab.BlTh & "'"
    
   Report7.RecordSelectionFormula = "{qryTopTw2Penyakit.tahun}=" & _
    rpt103BlThKab.pvNum(0).ValueInteger & ""
  
   Report7.PrinterSetup Me.hWnd
   CRViewer1.ReportSource = Report7
  
  Case 7
   Report8.FormulaFields(18).Text = "'" & rpt103BlThKab.BlTh & "'"
    
   Report8.RecordSelectionFormula = "{qryTopTw3Penyakit.tahun}=" & _
    rpt103BlThKab.pvNum(0).ValueInteger & ""
  
   Report8.PrinterSetup Me.hWnd
   CRViewer1.ReportSource = Report8
  
  Case 10
   Report9.FormulaFields(18).Text = "'" & rpt103BlThKab.BlTh & "'"
    
   Report9.RecordSelectionFormula = "{qryTopTw4Penyakit.tahun}=" & _
    rpt103BlThKab.pvNum(0).ValueInteger & ""
  
   Report9.PrinterSetup Me.hWnd
   CRViewer1.ReportSource = Report9
   
  End Select
 
 Case 9
   BlTh = "TAHUN " & Trim(Str(rpt10BlThKab.pvNum(0).ValueReal))
   Report10.FormulaFields(17).Text = "'" & BlTh & "'"
  
   Report10.RecordSelectionFormula = "{qryRekapLB1PerTahun.tahun}=" & _
     rpt10BlThKab.pvNum(0).ValueInteger
  
   Report10.PrinterSetup Me.hWnd
   CRViewer1.ReportSource = Report10
 
 Case 10
   BlTh = "TAHUN " & Trim(Str(rptRekapLB.pvNum(0).ValueReal))
   Report11.FormulaFields(16).Text = "'" & BlTh & "'"
  
   Report11.RecordSelectionFormula = "{qRekapLB1.tahun}=" & _
     rptRekapLB.pvNum(0).ValueInteger & " and {qRekapLB1.kdPenyakit}='" & _
     rptRekapLB.cKode & "'"
  
   Report11.PrinterSetup Me.hWnd
   CRViewer1.ReportSource = Report11
 
 Case 11
  BlTh = "PUSKESMAS " & Trim(rpt10BlThKabPus.txtdata(1).Text) & _
         " (" & UCase(rpt10BlThKabPus.cmbData(0).Text) & _
         " " & Trim(Str(rpt10BlThKabPus.pvNum(0).ValueReal)) & ")"
  Report12.FormulaFields(16).Text = "'" & BlTh & "'"
  
  Report12.RecordSelectionFormula = "{qrypTopPenyakit.bulan}= " & _
     rpt10BlThKabPus.cmbData(0).ListIndex + 1 & " and {qrypTopPenyakit.tahun}=" & _
     rpt10BlThKabPus.pvNum(0).ValueInteger & " and {qrypTopPenyakit.kdPuskesmas}='" & _
     Trim(rpt10BlThKabPus.txtdata(0).Text) & "'"
  
  Report12.PrinterSetup Me.hWnd
  CRViewer1.ReportSource = Report12
 
 Case 12
   BlTh = "PUSKESMAS " & Trim(rpt10BlThKabPus.txtdata(1).Text) & _
          " (TAHUN " & Trim(Str(rpt10BlThKabPus.pvNum(0).ValueReal)) & ")"
   Report13.FormulaFields(16).Text = "'" & BlTh & "'"
  
   Report13.RecordSelectionFormula = "{qryTopPThnPenyakit.tahun}=" & _
     rpt10BlThKabPus.pvNum(0).ValueInteger & " and {qryTopPThnPenyakit.kdPuskesmas}='" & _
     Trim(rpt10BlThKabPus.txtdata(0).Text) & "'"
  
   Report13.PrinterSetup Me.hWnd
   CRViewer1.ReportSource = Report13
 End Select
 
 CRViewer1.Zoom 100
 CRViewer1.ViewReport
 Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
 CRViewer1.Top = 0
 CRViewer1.Left = 0
 CRViewer1.Height = ScaleHeight
 CRViewer1.Width = ScaleWidth
End Sub
