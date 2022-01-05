' GLOBAL VARIABLES
Dim MainBookTitle As String
Dim BOOKMAP_ID As String
' ---------- ---------- ---------- ---------- ----------
Sub Main()

  Call Get_MainBookTitle
  Call Get_BOOKMAP_ID  
  Call Create_Bookmap
  Call Save_Bookmap
  
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
Sub Create_Bookmap()

  Dim MapFilename As String

  ' create new document
  Documents.Add
  
  ' initialise document
  ActiveDocument.PageSetup.Orientation = wdOrientLandscape
  ActiveDocument.PageSetup.TopMargin = 35
  ActiveDocument.PageSetup.RightMargin = 35
  ActiveDocument.PageSetup.BottomMargin = 35
  ActiveDocument.PageSetup.LeftMargin = 35
  Documents(1).Content.Font.Size = 9
  Documents(1).Content.Font.name = "Courier New"
  
  ' add header 'boilerplate'
  Documents(1).Content.InsertAfter "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine
  Documents(1).Content.InsertAfter "<!DOCTYPE bookmap PUBLIC ""-//OASIS//DTD DITA BookMap//EN"""
  Documents(1).Content.InsertAfter " ""bookmap.dtd"""
  Documents(1).Content.InsertAfter ">"
  Documents(1).Content.InsertAfter vbNewLine
  
  ' bookmap-element
  Documents(1).Content.InsertAfter "<bookmap id=""" & BOOKMAP_ID & """"
  Documents(1).Content.InsertAfter " xml:lang=""en_US" & """"
  Documents(1).Content.InsertAfter ">"
  Documents(1).Content.InsertAfter vbNewLine
  
  ' booktitle-element
  Documents(1).Content.InsertAfter vbNewLine
  Documents(1).Content.InsertAfter "  " & "<booktitle>" & vbNewLine
  Documents(1).Content.InsertAfter "    " & "<mainbooktitle>" & MainBookTitle & "</mainbooktitle>" & vbNewLine
  Documents(1).Content.InsertAfter "  " & "</booktitle>" & vbNewLine
  Documents(1).Content.InsertAfter vbNewLine
   
  ' add chapter-elements
  For Each Paragraph In Documents(2).Paragraphs
  
    If Paragraph.Range.Style = "dita-chapter" Then
      MapFilename = Left(Paragraph, Len(Paragraph) - 1)
      MapFilename = LCase(MapFilename)
      MapFilename = Replace(MapFilename, " ", "_")
      MapFilename = MapFilename & ".ditamap"
      
      ' chapter-element
      Documents(1).Content.InsertAfter "  <chapter href=""m_" & MapFilename & """"
      Documents(1).Content.InsertAfter " format=""ditamap" & """"
      Documents(1).Content.InsertAfter " scope=""local" & """"
      Documents(1).Content.InsertAfter " type=""map" & """"
      Documents(1).Content.InsertAfter " navtitle=""" & Left(Paragraph, Len(Paragraph) - 1) & " Map"""
      Documents(1).Content.InsertAfter "/>" & vbNewLine
    End If
  
  Next Paragraph
  
  ' add closing-bookmap element
  Documents(1).Content.InsertAfter "</bookmap>"
  
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
