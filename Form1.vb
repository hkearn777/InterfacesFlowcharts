Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Class Form1
  ' Create the Interfaces PUML commands to create a flowchart
  Dim ProgramVersion As String = "v0.0"
  ' Change-History.
  ' 2025-03-27 hk v0.0 - Initial version.
  '-----------------------------------------------------------------------------
  ' load the Excel References
  Dim objExcel As New Microsoft.Office.Interop.Excel.Application
  ' Interfaces spreadsheet 
  Dim workbook As Microsoft.Office.Interop.Excel.Workbook
  Dim FilesWorksheet As Microsoft.Office.Interop.Excel.Worksheet
  Dim theWorksheet As Microsoft.Office.Interop.Excel.Worksheet


  Dim DefaultFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault
  Dim SetAsReadOnly = Microsoft.Office.Interop.Excel.XlFileAccess.xlReadOnly

  Dim Delimiter As String = "|"

  ' define the lists of each of the interfaces
  'key source, value updn, type, status
  Dim internalDownstreamDict As New Dictionary(Of String, String)
  Dim externalDownstreamDict As New Dictionary(Of String, String)
  Dim internalUpstreamDict As New Dictionary(Of String, String)
  Dim externalUpstreamDict As New Dictionary(Of String, String)

  ' define the list of internal and external interfaces
  Dim internalInterfaces As New List(Of String)
  Dim externalInterfaces As New List(Of String)

  Dim swPuml As StreamWriter

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.Text = "InterfacesFlowcharts " & ProgramVersion
  End Sub
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub


  Private Sub btnGo_Click(sender As Object, e As EventArgs) Handles btnGo.Click
    ' ensure target application is entered
    If txtTargetApp.Text.Trim.Length = 0 Then
      MsgBox("Please enter the target application")
      Exit Sub
    End If
    ' ensure interface filename is entered
    If txtInterfaceFilename.Text.Trim.Length = 0 Then
      MsgBox("Please enter the interface filename")
      Exit Sub
    End If
    ' verify the interface file exists
    If Not File.Exists(txtInterfaceFilename.Text) Then
      MsgBox("The interface file does not exist")
      Exit Sub
    End If
    ' ensure puml folder name is entered
    If txtPumlFolderName.Text.Trim.Length = 0 Then
      MsgBox("Please enter the PUML folder name")
      Exit Sub
    End If
    ' ensure puml internal filename is entered
    If txtInternalPumlFilename.Text.Trim.Length = 0 Then
      MsgBox("Please enter the PUML internal filename")
      Exit Sub
    End If
    ' ensure puml external filename is entered
    If txtExternalPumlFilename.Text.Trim.Length = 0 Then
      MsgBox("Please enter the PUML external filename")
      Exit Sub
    End If


    Me.Cursor = Cursors.WaitCursor
    objExcel.Visible = False


    Dim InterfaceFileName As String = txtInterfaceFilename.Text
    workbook = objExcel.Workbooks.Open(InterfaceFileName, True, SetAsReadOnly)

    Call LoadAllLists()
    internalInterfaces = CombineInterfaceLists(internalDownstreamDict, internalUpstreamDict)
    externalInterfaces = CombineInterfaceLists(externalDownstreamDict, externalUpstreamDict)


    ' close the spreadsheet
    workbook.Close()
    objExcel.Quit()

    ' create the flowcharts
    Call CreateFlowchart("Internal", txtInternalPumlFilename.Text, internalInterfaces)
    Call CreateFlowchart("External", txtExternalPumlFilename.Text, externalInterfaces)



    Me.Cursor = Cursors.Default
    MessageBox.Show("Complete.")

  End Sub

  Sub LoadAllLists()
    ' using the spreadsheet to load all the external / download lists

    theWorksheet = workbook.Sheets.Item(1)
    theWorksheet.Activate()
    Dim MaxRows As Long = theWorksheet.UsedRange.Rows(theWorksheet.UsedRange.Rows.Count).row
    Dim MaxCols As Long = theWorksheet.UsedRange.Columns.Count
    Dim startRow As Integer = 2
    For row As Integer = startRow To MaxRows
      Dim sourceSystem As String = GrabExcelField(row, 2).
        ToUpper.Trim.Replace(" ", "_").
        Replace("-", "_").
        Replace("/", "_").
        Replace("(", "").
        Replace(")", "").
        Replace(",", "").
        Replace("__", "_")
      sourceSystem = sourceSystem.Replace("__", "_")

      Dim UpDn As String = GrabExcelField(row, 4).ToUpper.Trim
      Dim targetSystem As String = GrabExcelField(row, 5).
        ToUpper.Trim.Replace(" ", "_").
        Replace("-", "_").
        Replace("/", "_").
        Replace("(", "").
        Replace(")", "").
        Replace(",", "").
        Replace("__", "_")
      targetSystem = targetSystem.Replace("__", "_")

      Dim interfaceType As String = GrabExcelField(row, 9).ToUpper.Trim
      Dim reviewStatus As String = GrabExcelField(row, 11).ToUpper.Trim
      If reviewStatus.Contains("OBSOLETE??") Then
        Continue For
      End If

      ' decide which list to add to depending on the interface type and updn
      Dim theKey As String
      Dim theValue As String = interfaceType & Delimiter & reviewStatus
      Select Case interfaceType & "|" & UpDn
        Case "INTERNAL|DOWNSTREAM"
          theKey = targetSystem & Delimiter & UpDn
          If Not internalDownstreamDict.ContainsKey(theKey) Then
            internalDownstreamDict.Add(theKey, theValue)
          End If
        Case "EXTERNAL|DOWNSTREAM"
          theKey = targetSystem & Delimiter & UpDn
          If Not externalDownstreamDict.ContainsKey(theKey) Then
            externalDownstreamDict.Add(theKey, theValue)
          End If
        Case "INTERNAL|UPSTREAM"
          theKey = sourceSystem & Delimiter & UpDn
          If Not internalUpstreamDict.ContainsKey(theKey) Then
            internalUpstreamDict.Add(theKey, theValue)
          End If
        Case "EXTERNAL|UPSTREAM"
          theKey = sourceSystem & Delimiter & UpDn
          If Not externalUpstreamDict.ContainsKey(theKey) Then
            externalUpstreamDict.Add(theKey, theValue)
          End If
        Case Else
          MessageBox.Show("Unknown interface type:" & interfaceType & ",UpDn:" & UpDn & ",row:" & row)
      End Select

    Next

  End Sub
  Function GrabExcelField(ByRef theRow As Integer, ByRef theColumn As Integer) As String
    If theRow = 0 Then
      Return ""
    End If
    If theColumn = 0 Then
      Return ""
    End If
    Dim theValue As String = theWorksheet.Cells(theRow, theColumn).value2
    If theValue Is Nothing Then
      Return ""
    End If
    If theValue.Length = 0 Then
      Return ""
    End If
    Return theValue
  End Function

  Function CombineInterfaceLists(ByRef theDownstreamDict As Dictionary(Of String, String),
                           ByRef theUpstreamDict As Dictionary(Of String, String)) As List(Of String)

    Dim myInterfaceList As New List(Of String)

    Dim interfacesDict As New Dictionary(Of String, String)
    ' combine the internal lists
    For Each kvp As KeyValuePair(Of String, String) In theDownstreamDict
      If Not interfacesDict.ContainsKey(kvp.Key) Then
        interfacesDict.Add(kvp.Key, kvp.Value)
      End If
    Next
    For Each kvp As KeyValuePair(Of String, String) In theUpstreamDict
      If Not interfacesDict.ContainsKey(kvp.Key) Then
        interfacesDict.Add(kvp.Key, kvp.Value)
      End If
    Next


    ' sort interfaces dict by application (key)
    Dim sortedInterfacesDict = interfacesDict.OrderBy(Function(kvp) kvp.Key).ToList()

    ' go through each internal interface looking for both directions
    For x As Integer = 0 To sortedInterfacesDict.Count - 1
      If x = sortedInterfacesDict.Count - 1 Then  ' Ensure there is a next entry
        Exit For
      End If

      Dim sourceName(1) As String
      Dim upDn(1) As String
      Dim targetName(1) As String
      Dim interfaceType(1) As String
      Dim reviewStatus(1) As String

      ' Get fields from first entry
      Dim theKey As String() = sortedInterfacesDict(x).Key.Split(Delimiter)
      sourceName(0) = theKey(0)
      upDn(0) = theKey(1)
      Dim firstEntry As String() = sortedInterfacesDict(x).Value.Split(Delimiter)
      interfaceType(0) = firstEntry(0)
      reviewStatus(0) = firstEntry(1)

      ' Get fields from next entry
      theKey = sortedInterfacesDict(x + 1).Key.Split(Delimiter)
      sourceName(1) = theKey(0)
      upDn(1) = theKey(1)
      Dim secondEntry As String() = sortedInterfacesDict(x + 1).Value.Split(Delimiter)
      interfaceType(1) = secondEntry(0)
      reviewStatus(1) = secondEntry(1)

      Dim theValue As String
      If sourceName(0) = sourceName(1) Then
        theValue = sourceName(0) & Delimiter &
          "Both" & Delimiter &
          interfaceType(0) & "/" & interfaceType(1) & Delimiter &
          reviewStatus(0)
        x += 1
      Else
        theValue = sourceName(0) & Delimiter &
          upDn(0) & Delimiter &
          interfaceType(0) & "/" & Delimiter &
          reviewStatus(0)
      End If

      myInterfaceList.Add(theValue)
    Next
    Return myInterfaceList
  End Function

  Sub CreateFlowchart(ByVal myInterfaceLiteral As String,
                      ByRef myFilename As String,
                      ByRef myInterface As List(Of String))
    ' this routine will create the PUML files for the internal interfaces
    Dim theFilename As String = txtPumlFolderName.Text & myFilename
    swPuml = New StreamWriter(theFilename)

    swPuml.WriteLine("@startuml " & txtTargetApp.Text & " " & myInterfaceLiteral & " Interfaces")
    swPuml.WriteLine("header " & Me.Text & "(c), by IBM")
    swPuml.WriteLine("title " & txtTargetApp.Text & " " & myInterfaceLiteral & " Interfaces")
    swPuml.WriteLine("")

    ' write the applications/systems
    swPuml.WriteLine("rectangle " & txtTargetApp.Text)
    For Each theInterface As String In myInterface
      Dim theFields As String() = theInterface.Split(Delimiter)
      Dim theName As String = theFields(0)
      Dim theUpDn As String = theFields(1)
      Dim theType As String = theFields(2)
      Dim theStatus As String = theFields(3)
      swPuml.WriteLine("rectangle " & theName)
    Next
    swPuml.WriteLine("")

    Dim bothCnt As Integer = 0
    ' now we need to map the applications just the upstream to target application
    swPuml.WriteLine("' Upstream Apps to Target")
    For Each theInterface As String In myInterface
      Dim theFields As String() = theInterface.Split(Delimiter)
      Dim theName As String = theFields(0)
      Dim theUpDn As String = theFields(1)
      Dim theType As String = theFields(2)
      Dim theStatus As String = theFields(3)
      If theUpDn = "UPSTREAM" Then
        swPuml.WriteLine(theName & " -[#blue]down-> " & txtTargetApp.Text)
      End If
      If theUpDn = "Both" Then
        bothCnt += 1
      End If
    Next
    swPuml.WriteLine("")

    ' now we need to map the applications that are both up and down stream to/from target application
    swPuml.WriteLine("left to right direction")
    ' determine midpoint of both applications
    Dim midPoint As Integer = bothCnt \ 2
    Dim cnt As Integer = 0
    swPuml.WriteLine("' Upstream and Downstream Apps to Target")
    For Each theInterface As String In myInterface
      Dim theFields As String() = theInterface.Split(Delimiter)
      Dim theName As String = theFields(0)
      Dim theUpDn As String = theFields(1)
      Dim theType As String = theFields(2)
      Dim theStatus As String = theFields(3)
      If theUpDn = "Both" Then
        cnt += 1
        If cnt <= midPoint Then
          swPuml.WriteLine(txtTargetApp.Text & " <-[#red]> " & theName)
        Else
          swPuml.WriteLine(txtTargetApp.Text & " <-[#red]left-> " & theName)
        End If
      End If
    Next
    swPuml.WriteLine("")

    swPuml.WriteLine("top to bottom direction")
    ' now we need to map the applications that downstream to target application
    swPuml.WriteLine("' Downstream Apps to Target")
    For Each theInterface As String In myInterface
      Dim theFields As String() = theInterface.Split(Delimiter)
      Dim theName As String = theFields(0)
      Dim theUpDn As String = theFields(1)
      Dim theType As String = theFields(2)
      Dim theStatus As String = theFields(3)
      If theUpDn = "DOWNSTREAM" Then
        swPuml.WriteLine(txtTargetApp.Text & " -[#green]down-> " & theName)
      End If
    Next
    swPuml.WriteLine("")

    swPuml.WriteLine("Legend Bottom left")
    swPuml.WriteLine("    Blue Lines upstream")
    swPuml.WriteLine("    Green Lines downstream")
    swPuml.WriteLine("    Red Lines up/downstream")
    swPuml.WriteLine("endlegend")

    swPuml.WriteLine("")
    swPuml.WriteLine("@enduml")

    swPuml.Close()

  End Sub

End Class
