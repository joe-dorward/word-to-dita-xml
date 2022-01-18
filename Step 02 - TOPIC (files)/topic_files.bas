' GLOBAL VARIABLES
Dim Path As String
Dim MainBookTitle As String
Dim BOOKMAP_ID As String
Dim BookmapContent As String
' ---------- ---------- ---------- ---------- ----------
Sub Reload()
    ActiveDocument.Reload
End Sub
Sub Main()
  Call Reload
  Call Delete_ManualPageBreaks
  Call Add_NormalParagraphs

  Path = Get_Path

  Call Get_Headings
  Call Get_Topics
End Sub
Sub Delete_ManualPageBreaks()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^m"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Add_NormalParagraphs()
  ActiveDocument.Paragraphs.Add
  ActiveDocument.Paragraphs.Item(ActiveDocument.Paragraphs.Count).Range.Style = "Normal"
  ActiveDocument.Paragraphs.Add
  ActiveDocument.Paragraphs.Add
End Sub
Function Get_Path() As String

  Dim PathLength As Integer

  PathLength = Len(ActiveDocument.FullName) - Len(ActiveDocument.name)
  Get_Path = Left(ActiveDocument.FullName, PathLength)

End Function
Sub Get_Headings()

  Dim ParagraphNumber As Integer
  Dim RowNumber As Integer
  RowNumber = 1
  
  Dim TableNumber As Integer
  
  ' get paragraph count before the table is added
  Dim Paragraph_Count As Integer
  Paragraph_Count = ActiveDocument.Paragraphs.Count
  
  Selection.EndKey Unit:=wdStory ' go to end of document
  
  ' ADD META-DATA TABLE
  ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=1, NumColumns:= _
    4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed
  
  ' get table count
  TableNumber = ActiveDocument.Tables.Count
  
  ActiveDocument.Tables.Item(TableNumber).Borders.InsideLineStyle = wdLineStyleSingle
  ActiveDocument.Tables.Item(TableNumber).Borders.OutsideLineStyle = wdLineStyleSingle
  
  ' format table
  ActiveDocument.Tables.Item(TableNumber).Range.Font.Size = 9
  
  ' header row text
  ActiveDocument.Tables.Item(TableNumber).Cell(1, 1).Range.Text = "Text"
  ActiveDocument.Tables.Item(TableNumber).Cell(1, 2).Range.Text = "Style"
  ActiveDocument.Tables.Item(TableNumber).Cell(1, 3).Range.Text = "P.Start"
  ActiveDocument.Tables.Item(TableNumber).Cell(1, 4).Range.Text = "P.End"

  ' get headings
  For Each Paragraph In ActiveDocument.Paragraphs
    ParagraphNumber = ParagraphNumber + 1

    ' get from
    If (Paragraph.Range.Style = "Heading 1") Then

      ActiveDocument.Tables.Item(TableNumber).Rows.Add
      RowNumber = RowNumber + 1
      
      ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 1).Range.Text = Left(Paragraph.Range.Text, Len(Paragraph.Range.Text) - 1)
      ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 2).Range.Text = Paragraph.Range.Style
      ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 3).Range.Text = ParagraphNumber ' Start
      
      If (RowNumber > 2) Then ' not first heading
        ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber - 1, 4).Range.Text = ParagraphNumber - 1 ' End
      End If
      
    End If
    
  Next Paragraph

  ActiveDocument.Tables.Item(TableNumber).Columns.AutoFit
  
  ' put the paragraph count (before the table was added) into the last 'paragraph end' cell
  ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 4).Range.Text = Paragraph_Count - 2
  
End Sub
Sub Get_Topics()

  Dim TableNumber As Integer
  TableNumber = ActiveDocument.Tables.Count
  
  Dim RowNumber As Integer ' row in the table
  
  Dim CellText As String
  Dim CellText_Length As Integer
  
  Dim TopicStart As Integer
  Dim TopicEnd As Integer
  Dim TopicContent As Range
  Dim Topic_Document_Filename As String
  
  For RowNumber = 2 To ActiveDocument.Tables.Item(TableNumber).Rows.Count
    ' MsgBox ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 3)
    Topic_Document_Filename = ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 1).Range.Text
    Topic_Document_Filename = Left(Topic_Document_Filename, Len(Topic_Document_Filename) - 2)
    Topic_Document_Filename = "(" & (RowNumber - 1) & ") " & Topic_Document_Filename

    ' get start paragraph number
    CellText = ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 3).Range.Text
    CellText_Length = Len(ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 3).Range.Text) - 2
    CellText = Left(CellText, CellText_Length)
    'MsgBox CellText
    Set TopicContent = ActiveDocument.Paragraphs(CellText).Range
  
    ' get end paragraph number
    CellText = ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 4).Range.Text
    CellText_Length = Len(ActiveDocument.Tables.Item(TableNumber).Cell(RowNumber, 4).Range.Text) - 2
    CellText = Left(CellText, CellText_Length)
    'MsgBox CellText
    TopicContent.End = ActiveDocument.Paragraphs(CellText).Range.End

    ' create new document
    TopicContent.Select
    Selection.Copy
    Documents.Add
    Selection.PasteAndFormat (wdUseDestinationStylesRecovery)

    ' save new document
    Dim TheFileDialog As FileDialog
    Set TheFileDialog = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
  
    TheFileDialog.InitialFileName = Path & Topic_Document_Filename
    TheFileDialog.Show
    TheFileDialog.Execute
    
    ' close the new document
    ActiveDocument.Close
  Next
  
  Selection.Collapse
  
End Sub
