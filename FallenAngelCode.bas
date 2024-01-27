Attribute VB_Name = "Solution"
' use this module to write your solution to the homework assignment

' Determines in a certain stock is considered a "Fallen Angel" based on historic and current PE Ratios
Sub processFallenAngel()

    ' Select the "Fallen Angel" sheet
    Sheets("Fallen Angel").Select
    
   ' Start in cell B4 (the stock ticker symbols are located contiguously below cell B3)
    Range("B4").Select
    
  ' Loop until an empty cell is encountered in column B
  Do Until ActiveCell.Value = ""
    Call getFallenAngelData
  Loop

  
End Sub


' Retrieves the PE Ratio data and determines a stock's Fallen Angel status
Sub getFallenAngelData()

    ' Copy the ticker from the active cell
    Selection.Copy
    
    ' Switch to the "WebQuery" sheet
    Sheets("WebQuery").Select
    
    ' Paste the ticker in cell A1
    Range("A1").Select
    ActiveSheet.Paste
    
    ' Select cell A2 and clear clipboard
    Range("A2").Select
    Application.CutCopyMode = False
    
    ' Configure the web query with the URL based on the ticker
    With Selection.QueryTable
        .Connection = "URL;https://ycharts.com/companies/" & Range("A1").Value & "/pe_ratio"
        
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With


    ' Declare variables for extracting information from the web query result
    Dim targetCell As Range
    Dim startPosition As Integer
    Dim extractedText As String
    
    'Find the cell containing the text "PE Ratio: "
    Set targetCell = Cells.Find(What:="PE Ratio: ", After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
        
    ' Check if "PE Ratio: " is found
    If Not targetCell Is Nothing Then
        ' Get the starting position of the text within the cell
        startPosition = InStr(1, targetCell.Value, "PE Ratio: ")
        
        ' Extract the 5 characters after the search text
        extractedText = Mid(targetCell.Value, startPosition + Len("PE Ratio: "), 5)
        
    Else
        ' Display a message if the search text is not found
        MsgBox "Search text not found."
    End If
    
    ' Switch back to the "Fallen Angel" sheet
    Sheets("Fallen Angel").Select
    
    ' Paste the extracted PE ratio in the current cell
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.Value = extractedText
    
    ' Copy the 5-year high PE ratio from the web query result
        Sheets("WebQuery").Select
    Cells.Find(What:="Maximum", After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Activate
    
    ' Copy the value from the cell above the found cell
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.Copy
    
    ' Switch back to the "Fallen Angel" sheet
    Sheets("Fallen Angel").Select
    
    ' Paste the 5-year high PE ratio in the next column
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    
    ' Determine if the stock is a fallen angel and update the status in the next column
    If (((ActiveCell.Offset(0, 0).Range("A1").Value) / 2) > ActiveCell.Offset(0, -1).Range("A1").Value) Then
        ActiveCell.Offset(0, 1).Range("A1").Value = "Yes"
    Else
        ActiveCell.Offset(0, 1).Range("A1").Value = "No"
    End If
    
    ' Move to the next cell down to avoid an infinite loop
    ActiveCell.Offset(1, -2).Range("A1").Select


End Sub


Sub runButton(control As IRibbonControl)
  ' this procedures allows the processFallenAngel procedure to be invoked from the "assignment" ribbon

  Application.Run (control.ID)
End Sub

Sub ClearResults()
  ThisWorkbook.Activate
  Sheets("Fallen Angel").Activate
  Range(Sheets("Fallen Angel").Range("c4"), Sheets("Fallen Angel").Cells(Rows.Count, 2).End(xlUp).Offset(0, 3)).ClearContents
End Sub
