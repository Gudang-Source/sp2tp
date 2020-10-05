Attribute VB_Name = "Modulmain"
Public kdProp As String, nmProp As String
Public kdKab As String, nmKab As String
Public kdKec As String, nmKec As String
Public kdPus As String, nmPus As String
Public appTipe As String, almt As String

Global Scatter_Code As Variant, Scatter_Code1 As Variant
Global Scatter_Code2 As Variant, Scatter_Code3 As Variant
Global Scatter_Code4 As Variant, Scatter_Code5 As Variant
Global Field_No As Integer, Field_No1 As Integer
Global Field_No2 As Integer, Field_No3 As Integer
Global Field_No4 As Integer, noLap As Integer
Global CnString As String, RsString As String

Option Explicit

Sub Main()
 Dim adoRec As ADODB.Recordset
 Dim adoCon As ADODB.Connection
 Dim strSQL As String
 
 Set adoCon = New ADODB.Connection
 adoCon.CursorLocation = adUseClient
 'adoCon.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
 '            "Data Source=" & App.Path & "\dinkes07.mdb;"
 adoCon.Open "DSN=dinkesLab"
 strSQL = "select * from tbSetGlobal"
 Set adoRec = New ADODB.Recordset
   adoRec.Open strSQL, adoCon, adOpenStatic, adLockOptimistic
 If adoRec.EOF And adoRec.BOF Then
  Load frmGlobal
  frmGlobal.Show 1
 Else
  kdProp = adoRec.Fields("kdProp").Value
  nmProp = adoRec.Fields("nmProp").Value
  kdKab = adoRec.Fields("kdKab").Value
  nmKab = adoRec.Fields("nmKab").Value
  kdKec = adoRec.Fields("kdKec").Value
  nmKec = adoRec.Fields("nmKec").Value
  kdPus = adoRec.Fields("kdPus").Value
  nmPus = adoRec.Fields("nmPus").Value
  appTipe = adoRec.Fields("vCheck4").Value
  almt = adoRec.Fields("alamat_instansi").Value
 End If
 adoRec.Requery
 If adoRec.RecordCount > 0 Then
  'Splash Screen (Under Construction)
  Load frmLogin
  frmLogin.Show
 Else
  MsgBox "Anda Harus Melakukan Setting Aplikasi", vbOKOnly + vbInformation, "DINKES - SP2TP"
  End
 End If
End Sub

Public Sub Save_Code(vString As Variant, vString1 As Variant, _
         vString2 As Variant, vString3 As Variant, _
         vString4 As Variant, vString5 As Variant)
    Scatter_Code = vString
    Scatter_Code1 = vString1
    Scatter_Code2 = vString2
    Scatter_Code3 = vString3
    Scatter_Code4 = vString4
    Scatter_Code5 = vString5
End Sub

Public Sub ShowFind(vCn As String, vRs As String, _
        Optional nField As Integer, Optional nField1 As Integer, _
        Optional nField2 As Integer, Optional nField3 As Integer, _
        Optional nField4 As Integer)
    CnString = vCn
    RsString = vRs
    Field_No = nField
    Field_No1 = nField1
    Field_No2 = nField2
    Field_No3 = nField3
    Field_No4 = nField4
    
    FrmScatter.Show 1
End Sub


Public Sub DatagridColumnAutoResize(ByRef oDataGrid As DataGrid, _
    ByRef oForm As Form)
Dim i As Integer, iMax As Integer
Dim t As Integer, tMax As Integer
Dim iWidth As Integer
Dim vBMark As Variant
Dim aWidth As Variant
Dim cText As String
Dim oFont As Font

'    On Error Resume Next

    'need this to make TextWidth()
    'work with prossibly different font in DG
    Set oFont = oForm.Font
    oForm.Font = oDataGrid.Font

    iMax = oDataGrid.Columns.count - 1
    ReDim aWidth(iMax)

    For i = 0 To iMax   'init maxwidth holder
       aWidth(i) = 0
    Next

    'one visible page to get to an estimate
    tMax = oDataGrid.VisibleRows - 1
    If tMax > 0 Then
        For t = 0 To tMax   'number of rows
            vBMark = oDataGrid.GetBookmark(t)
            For i = 0 To iMax   'number of columns
                cText = oDataGrid.Columns(i).CellText(vBMark)
                iWidth = oForm.TextWidth(cText)
                If iWidth + ((12 * Len(cText)) + 220) > aWidth(i) Then
                    'the font is right, the stringlength too, but
                    'still some misalignment on long stings. So we
                    'have to fiddle this a bit by hand...
                    aWidth(i) = iWidth + ((12 * Len(cText)) + 220)
                End If
                If t = 0 Then   'take care of the headers
                    iWidth = oForm.TextWidth(oDataGrid.Columns( _
                        i).Caption)
                    If iWidth + ((12 * Len(cText)) + 220) > aWidth( _
                        i) Then
                        aWidth(i) = iWidth + ((12 * Len(cText)) + 220)
                    End If
                End If
            Next
        Next
        For i = 0 To iMax   ' finally set the new column width
            oDataGrid.Columns(i).Width = aWidth(i)
        Next
    End If
    oForm.Font = oFont
End Sub

Public Function Export2XL(InitRow As Long, DBAccess As String, DBTable As String, nmFile As String, dir As String, frm As Form) As Long

Dim cn As New ADODB.Connection           'Use for the connection string
Dim cmd As New ADODB.Command          'Use for the command for the DB
Dim Rs As New ADODB.Recordset             'Recordset return from the DB
Dim MyIndex As Integer                            'Used for Index
Dim MyRecordCount As Long                    'Store the number of record on the table
Dim MyFieldCount As Integer                    'Store the number of fields or column
Dim ApExcel As Object                             'To open Excel
Dim xlBook As Object
Dim xlSheet As Object
Dim MyCol As String
Dim Response As Integer

Set ApExcel = CreateObject("Excel.application")  'Creates an object
Set xlBook = ApExcel.Workbooks.Add                                   'Adds a new book.
Set xlSheet = ApExcel.Worksheets.Add
xlSheet.Name = nmFile

'Set the connection string
'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBAccess
cn.ConnectionString = "DSN=dinkesLab"
'Open the connection
cn.Open

'Check that the connection is open
If cn.State = 0 Then cn.Open
Set cmd.ActiveConnection = cn
cmd.CommandText = DBTable
cmd.CommandType = adCmdText
Set Rs = cmd.Execute
'Count the number of fields or column
MyFieldCount = Rs.Fields.count

'Fill the first line with the name of the fields
For MyIndex = 0 To MyFieldCount - 1
    xlSheet.Cells(InitRow, (MyIndex + 1)).Formula = Rs.Fields(MyIndex).Name   'Write Title to a Cell
    xlSheet.Cells(InitRow, (MyIndex + 1)).Font.Bold = True
    xlSheet.Cells(InitRow, (MyIndex + 1)).Interior.ColorIndex = 36
    xlSheet.Cells(InitRow, (MyIndex + 1)).WrapText = True
Next

'Draw border on the title line
MyCol = Chr((64 + MyIndex)) & InitRow
xlSheet.Range("A" & InitRow & ":" & MyCol).Borders.Color = RGB(0, 0, 0)
MyRecordCount = 1 + InitRow

'Fill the excel book with the values from the database
frm.pgb.Min = 0
frm.pgb.Value = 0
Do While Rs.EOF = False
For MyIndex = 1 To MyFieldCount
    xlSheet.Cells(MyRecordCount, MyIndex).Formula = Rs((MyIndex - 1)).Value     'Write Value to a Cell
    xlSheet.Cells(MyRecordCount, MyIndex).WrapText = False 'Format the Cell
Next
    MyRecordCount = MyRecordCount + 1
    frm.pgb.Max = MyRecordCount
    frm.pgb.Value = frm.pgb.Max
    Delay 6000
    Rs.MoveNext
    'If MyRecordCount > 50 Then
    '    Exit Do
    'End If
Loop

'Close the connection with the DB
Rs.Close

'Return the last position in the workbook
Export2XL = MyRecordCount

'xlBook.SaveAs App.Path & "\" & nmFile
xlBook.SaveAs dir
'ApExcel.Visible = True                                       'This enable you to see the process in Excel

'Information that process has been completed
Response = MsgBox("Proses Export Selesai", vbOKOnly + vbInformation, "DINKES")

Set xlSheet = Nothing
Set xlBook = Nothing
Set ApExcel = Nothing
Excel.Application.Quit

End Function

Public Sub Delay(Waktu As Long)
 Dim i As Long
 For i = 0 To Waktu
 Next
End Sub

