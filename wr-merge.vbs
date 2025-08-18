option explicit

sub main

  dim xl, pt
  dim xl_file, pt_file
  dim oSlide, oShape, oWorkbook, oWorksheet, oTable, iRow, iCol
  dim i, j
  
  process_delete "EXCEL.EXE"
  process_delete "POWERPNT.EXE"

  console_debug "start"
  xl_file = getcwd & "\merge.xlsx"
  pt_file = getcwd & "\merge.pptx"
  
  set xl = CreateObject("Excel.Application")
  xl.Visible = true

  set pt = CreateObject("PowerPoint.Application")
  pt.Visible = true
  
  xl.Workbooks.Open(xl_file)
  set oWorkbook = xl.Workbooks.Item(1)

  pt.Presentations.Open(pt_file)
  
  for i = 1 to pt.Presentations(1).Slides.Count
    console_debug "processing Slide " & i
    set oSlide = pt.Presentations(1).Slides(i)
    oSlide.Select
    For j = 1 to oSlide.Shapes.Count
      set oShape = oSlide.Shapes.Item(j)
      If oShape.HasTextFrame Then
        console_debug "processing Slide " & i & " Shape " & j & " TextFrame"
        If oShape.TextFrame.HasText Then
          setInfoValue oWorkbook, oShape.TextFrame.TextRange
        End If
        setPictureValue oWorkbook, oSlide, oShape
      ElseIf oShape.HasTable Then
        console_debug "processing Slide " & i & " Shape " & j & " Table"
        Set oTable = oShape.Table
        For iRow = 1 To oTable.Rows.Count
          For iCol = 1 To oTable.Columns.Count
            setInfoValue oWorkbook, oTable.Cell(iRow, iCol).Shape.TextFrame.TextRange
          Next
        Next
      End If
    Next
  Next
  
  delPicturePlaceholders pt.Presentations.Item(1)
  
  console_debug "close Workbook"
  oWorkbook.Close
  console_debug "close Excel"
  xl.Quit
  console_debug "terminating .."
  WScript.Sleep 4000
'  console_debug "close PowerPoint"
'  pt.Quit
  console_debug "completed"
end sub

Sub setInfoValue(oWorkbook, oTextRange)

  Dim oRegExp
  Dim oMatches, oMatch
  Dim sReplace
  dim sWorksheet
  
  Set oRegExp = New RegExp
  oRegExp.Pattern = "\${([A-Za-z0-9]+)\.([A-Z]+[0-9]+)}"
  oRegExp.Global = True
  Set oMatches = oRegExp.Execute(oTextRange.Text)
  For Each oMatch In oMatches
    ' console_debug oMatch.Value
    sWorksheet = oMatch.SubMatches.Item(0)
    sReplace = getInfoCellValue2(oWorkbook, sWorksheet, oMatch.SubMatches.Item(1))
    oTextRange.Replace oMatch.Value, sReplace
  Next
  
End Sub

' Get cell value from sheet

Function getInfoCellValue2(oWorkbook, sWorksheet, CellRef)
  dim oWorksheet
  set oWorksheet = get_worksheet(oWorkbook, sWorksheet)
  getInfoCellValue2 = oWorksheet.Range(CellRef).Text
End Function

Sub setPictureValue(oWorkbook, oSlide, oShape)

  Dim oRegExp
  Dim oMatches, oMatch, oPicture
  Dim oWorksheet
  
  Set oRegExp = New RegExp
    
  oRegExp.Pattern = "\${([A-Za-z0-9]+)\.([A-Z]+[0-9]+:[A-Z]+[0-9]+)}"
  oRegExp.Global = True
  Set oMatches = oRegExp.Execute(oShape.TextFrame.TextRange.Text)
  If oMatches.Count = 0 Then
    Exit Sub
  End If
  
  For Each oMatch In oMatches
    set oWorksheet = get_worksheet(oWorkbook, oMatch.SubMatches.Item(0))
    oWorksheet.Range(oMatch.SubMatches.Item(1)).CopyPicture
    ' 0 - ppPasteDefault
    Set oPicture = oSlide.Shapes.PasteSpecial(0)
    ' 0 - msoFalse
    oPicture.LockAspectRatio = 0
    oPicture.Left = oShape.Left
    oPicture.Top = oShape.Top
    oPicture.Width = oShape.Width
    oPicture.Height = oShape.Height
  Next
End Sub

' Delete picture placeholders (Shape with TextFrame)
' @param Presentation p

Sub delPicturePlaceholders(p)

  console_debug "delete picture placeholders .."
  
  Dim oRegExp 
  Dim oMatches, oMatch
  Dim oSlide, oShape
  Dim i 
  
  Set oRegExp = New RegExp
  
  oRegExp.Pattern = "\${([A-Za-z0-9]+)\.([A-Z]+[0-9]+:[A-Z]+[0-9]+)}"
  oRegExp.Global = True
  
  For Each oSlide In p.Slides
    For i = oSlide.Shapes.Count To 1 Step -1
      Set oShape = oSlide.Shapes(i)
      If oShape.HasTextFrame Then
        Set oMatches = oRegExp.Execute(oShape.TextFrame.TextRange.Text)
        If oMatches.Count > 0 Then
          oShape.Delete
        End If
      End If
    Next
  Next
End Sub

' Get worksheet from a workbook by name
' "F" -> "FLPP"

function get_worksheet(oWorkbook, sPattern)

  dim oWorksheet
  dim i

  for i = 1 to oWorkbook.Worksheets.Count
    set oWorksheet = oWorkbook.Worksheets.Item(i)
    if mid(oWorksheet.Name, 1 , len(sPattern)) = sPattern then
      set get_worksheet = oWorksheet
      exit function
    end if
  next

  console_error "Worksheet not found. " & sPattern

end function

sub process_delete(process_name)
  dim oShell
  set oShell = CreateObject("WScript.Shell")
  oShell.Run "taskkill /f /im " & process_name, 0, true
end sub

' @link https://stackoverflow.com/a/21225466/82126

function getcwd
  getcwd = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)
end function

sub console_debug(message)
  WScript.Echo "DEBUG : " & message
end sub

sub console_error(message)
  WScript.Echo "ERROR : " & message
  WScript.Quit
end sub

main
