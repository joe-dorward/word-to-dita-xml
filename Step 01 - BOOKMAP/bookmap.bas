' GLOBAL VARIABLES
Dim MainBookTitle As String
Dim BOOKMAP_ID As String
Dim BookmapContent As String
' ---------- ---------- ---------- ---------- ----------
Sub Main()

  BookmapContent = ""
  
  Call Delete_ManualPageBreaks
  
  Call Get_MainBookTitle
  Call Get_BOOKMAP_ID
  
  Call Add_BookmapHeader
  Call Add_BookmapElement
  Call Add_BooktitleElement
  Call Add_ChapterElements
    
  Call Create_Bookmap
  Call Save_Bookmap
  
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
Sub Get_MainBookTitle()
  ' you may need to customise this sub-procedure further

  ' get the word document filename
  MainBookTitle = ActiveDocument.name
  
  ' strip off '.docx' extension
  MainBookTitle = Replace(MainBookTitle, ".docx", "")
  
  ' replace any underscores with spaces
  MainBookTitle = Replace(MainBookTitle, "_", " ")

End Sub
Sub Get_BOOKMAP_ID()
  ' you may need to customise this sub-procedure further

  ' get bookmap id from 'MainBookTitle'
  BOOKMAP_ID = "b_" & LCase(MainBookTitle)
  
  ' replace any spaces with underscores
  BOOKMAP_ID = Replace(BOOKMAP_ID, " ", "_")
  
  ' remove any hyphen-underscores
  BOOKMAP_ID = Replace(BOOKMAP_ID, "-_", "")
  
  ' replace any hyphens with underscores
  BOOKMAP_ID = Replace(BOOKMAP_ID, "-", "_")
  
  ' replace any dots with underscores
  BOOKMAP_ID = Replace(BOOKMAP_ID, ".", "_")
  
End Sub
Sub Add_BookmapHeader()
  BookmapContent = BookmapContent & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine
  BookmapContent = BookmapContent & "<!DOCTYPE bookmap PUBLIC ""-//OASIS//DTD DITA BookMap//EN"""
  BookmapContent = BookmapContent & " ""bookmap.dtd"""
  BookmapContent = BookmapContent & ">"
  BookmapContent = BookmapContent & vbNewLine
End Sub
Sub Add_BookmapElement()
  BookmapContent = BookmapContent & "<bookmap id=""" & BOOKMAP_ID & """"
  BookmapContent = BookmapContent & " xml:lang=""en_US" & """"
  BookmapContent = BookmapContent & ">"
  BookmapContent = BookmapContent & vbNewLine
End Sub
Sub Add_BooktitleElement()
  BookmapContent = BookmapContent & vbNewLine
  BookmapContent = BookmapContent & "  " & "<booktitle>" & vbNewLine
  BookmapContent = BookmapContent & "    " & "<mainbooktitle>" & MainBookTitle & "</mainbooktitle>" & vbNewLine
  BookmapContent = BookmapContent & "  " & "</booktitle>" & vbNewLine
  BookmapContent = BookmapContent & vbNewLine
End Sub
Sub Add_ChapterElements()

  For Each Paragraph In ActiveDocument.Paragraphs
  
    ' LOOKS FOR WORD STYLE BY NAME
    If Paragraph.Range.Style = "Heading 1" Then
      MapFilename = Left(Paragraph, Len(Paragraph) - 1)
      MapFilename = LCase(MapFilename)
      MapFilename = Replace(MapFilename, " ", "_")
      MapFilename = MapFilename & ".ditamap"
      
      ' chapter-element
      BookmapContent = BookmapContent & "  <chapter href=""m_" & MapFilename & """"
      BookmapContent = BookmapContent & " format=""ditamap" & """"
      BookmapContent = BookmapContent & " scope=""local" & """"
      BookmapContent = BookmapContent & " type=""map" & """"
      BookmapContent = BookmapContent & " navtitle=""" & Left(Paragraph, Len(Paragraph) - 1) & " Map"""
      BookmapContent = BookmapContent & "/>" & vbNewLine
    End If
    
  Next Paragraph
  
  ' add closing-bookmap element
  BookmapContent = BookmapContent & "</bookmap>"
End Sub
Sub Create_Bookmap()

  Dim MapFilename As String

  ' add new document
  Documents.Add

  ' add bookmap content
  ActiveDocument.Content.Text = BookmapContent
  
  ' initialise document
  ActiveDocument.PageSetup.Orientation = wdOrientLandscape
  ActiveDocument.PageSetup.TopMargin = 35
  ActiveDocument.PageSetup.RightMargin = 35
  ActiveDocument.PageSetup.BottomMargin = 35
  ActiveDocument.PageSetup.LeftMargin = 35
  
  ActiveDocument.Content.Font.Size = 9
  ActiveDocument.Content.Font.name = "Courier New"
  
  ActiveDocument.ActiveWindow.View.Type = wdPrintView
  ActiveDocument.ActiveWindow.View.Zoom.Percentage = 100
  
End Sub
Sub Save_Bookmap()
  ' save-as text-file
  ' change filename-extension to 'ditamap'

  Dim PathLength As Integer
  Dim Path As String
    
  Dim TheFileDialog As FileDialog
  Set TheFileDialog = Application.FileDialog(FileDialogType:=msoFileDialogSaveAs)
  
  ' get path
  PathLength = Len(Documents(2).FullName) - Len(Documents(2).name)
  Path = Left(Documents(2).FullName, PathLength)
    
  TheFileDialog.InitialFileName = Path & BOOKMAP_ID
  TheFileDialog.Show
  TheFileDialog.Execute

End Sub
