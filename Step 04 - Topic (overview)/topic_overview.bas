' GLOBAL VARIABLES
Dim TopicTitle As String
Dim TOPIC_ID As String
' ---------- ---------- ---------- ---------- ----------
Sub Reload()
    ActiveDocument.Reload
End Sub
Sub main()

  Call ActiveDocument.Reload
  Call Add_Style_Done
  TopicTitle = Get_TopicTitle
  Call Get_TOPIC_ID

  Call Add_UnformattedLastCharacter ' to each paragraph
  Call LessThan_To_LT
  Call GreaterThan_To_GT
  
  Call Convert_InlineFormatting
  Call Convert_Bullets_To_ListItems
  
  Call Add_TopicHeader
  Call Add_TopicFooter
    
  MsgBox "Done"
End Sub
' =============== =============== =============== =============== ===============
Sub Add_Style_Done()
    
    On Error GoTo Add_Style
        If ActiveDocument.Styles("Done").InUse Then
            Exit Sub
        End If

Add_Style:
    ActiveDocument.Styles.Add name:="Done", Type:=wdStyleTypeParagraph
    With ActiveDocument.Styles("Done").Font
        .name = "Courier New"
        .Size = 10
        .Bold = False
        .Italic = False
        .Color = RGB(0, 175, 75)
        '.Fill =
    End With
End Sub
Sub Test_Get_TopicTitle()
  TopicTitle = Get_TopicTitle
End Sub
Function Get_TopicTitle()
  'MsgBox Left(ActiveDocument.Paragraphs.Item(1).Range.Text, Len(ActiveDocument.Paragraphs.Item(1).Range.Text) - 1)
  
  Dim TopicTitle As String
  
  ActiveDocument.Paragraphs.Item(1).Range.Style = "Normal"
    
  TopicTitle = ActiveDocument.Paragraphs.Item(1).Range.Text
  TopicTitle = Left(TopicTitle, Len(TopicTitle) - 1)
  
  Get_TopicTitle = TopicTitle
End Function
Sub Get_TOPIC_ID()
  ' do not make this a function

  TOPIC_ID = "t_" & LCase(TopicTitle)
  TOPIC_ID = Replace(TOPIC_ID, " ", "_")
  TOPIC_ID = Replace(TOPIC_ID, "-_", "")
  TOPIC_ID = Replace(TOPIC_ID, "-", "_")
  TOPIC_ID = Replace(TOPIC_ID, ".", "_")
  
End Sub
Sub Add_UnformattedLastCharacter()
  ' ensure that the last-character of a prargraph is unformatted

  Dim NormalFontName As String
  Dim ParagraphLength As Integer
  
  NormalFontName = ActiveDocument.Styles("Normal").Font.name

  For Each Paragraph In ActiveDocument.Paragraphs
    ParagraphLength = Len(Paragraph.Range.Text) - 1 ' don't count paragraph mark
      
    If (ParagraphLength > 0) Then
      'Paragraph.Range.Characters.Item(ParagraphLength).Select
      Paragraph.Range.Characters.Item(ParagraphLength).InsertAfter " "
      Paragraph.Range.Characters.Item(ParagraphLength + 1).Font.name = NormalFontName
    End If
      
  Next Paragraph
End Sub
Sub LessThan_To_LT()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<"
        .Replacement.Text = "&lt;"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub GreaterThan_To_GT()
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ">"
        .Replacement.Text = "&gt;"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub
Sub Convert_InlineFormatting()
 
  Dim GotStart As Boolean
  Dim GotEnd As Boolean
  Dim CharacterCounter As Integer
  
  GotStart = False
  GotEnd = False
  CharacterCounter = 2
    
  For Each Character In ActiveDocument.Characters
    'MsgBox Character
    
    If CharacterCounter < ActiveDocument.Characters.Count - 2 Then
    
      ' GotStart
      If (ActiveDocument.Characters.Item(CharacterCounter).Font.name = "Courier New") And _
        Not (ActiveDocument.Characters.Item(CharacterCounter - 1).Font.name = "Courier New") Then
          GotStart = True
      End If
      
      If GotStart Then
        Set MyRange = ActiveDocument.Range(CharacterCounter - 1, CharacterCounter)
        MyRange.Select
        Selection.TypeText Text:="<tt>" & Selection
        
        'MsgBox "OK"
        GotStart = False
      End If
      
      ' GotEnd
      If Not (ActiveDocument.Characters.Item(CharacterCounter).Font.name = "Courier New") And _
         (ActiveDocument.Characters.Item(CharacterCounter - 1).Font.name = "Courier New") Then
        GotEnd = True
      End If
      
      If GotEnd Then
        Set MyRange = ActiveDocument.Range(CharacterCounter - 1, CharacterCounter)
        MyRange.Select
        Selection.TypeText Text:="</tt>" & Selection
  
        'MsgBox "OK"
        GotEnd = False
      End If
      
      CharacterCounter = CharacterCounter + 1
    Else
      'MsgBox "Woops"
      Exit Sub
    End If
    
  Next Character
  
  'MsgBox CharacterCounter
End Sub
Sub Convert_Bullets_To_ListItems()
  
    For Each Paragraph In ActiveDocument.Paragraphs
    
      If Paragraph.Range.ListFormat.ListType = wdListBullet Then
        
        Paragraph.Range.Select
        
        Selection.HomeKey Unit:=wdLine
        Selection.TypeText Text:="    <li>"
        Selection.EndKey Unit:=wdLine
        Selection.TypeText Text:="</li>"
          
        Paragraph.Range.ListFormat.RemoveNumbers
        Paragraph.Style = "Done"
      End If
          
    Next Paragraph

End Sub
Sub Add_TopicHeader()

  ActiveDocument.Paragraphs.Item(1).Range.Select
  Selection.Style = "Done"

  ' add header 'boilerplate'
  Selection = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine
  Selection = Selection.Text & "<!DOCTYPE topic PUBLIC ""-//OASIS//DTD DITA Topic//EN"""
  Selection = Selection.Text & " ""topic.dtd"""
  Selection = Selection.Text & ">"
  Selection = Selection.Text & vbNewLine
  
    ' topic-element
  Selection = Selection.Text & "<topic id=""" & TOPIC_ID & """"
  Selection = Selection.Text & " xml:lang=""en_US" & """"
  Selection = Selection.Text & ">"
  Selection = Selection.Text & vbNewLine
  
  ' title-element
  Selection = Selection.Text & "  " & "<title>" & TopicTitle & "</title>" & vbNewLine
  Selection = Selection.Text & vbNewLine
  
  ' body-element
  Selection = Selection.Text & "  " & "<body>" & vbNewLine
  Selection.Collapse
  
End Sub
Sub Add_TopicFooter()
  ActiveDocument.Paragraphs.Item(ActiveDocument.Paragraphs.Count).Range.Select
  Selection.Style = "Done"
  Selection.Collapse
  
    ' close body-element
  ActiveDocument.Content.InsertAfter vbNewLine
  ActiveDocument.Content.InsertAfter "  " & "</body>" & vbNewLine
  ActiveDocument.Content.InsertAfter "</topic>" & vbNewLine
  
End Sub
Sub Make_Selected_SubList()

  Dim SubList As String
  Dim ListItem As String
  
  Selection.Style = "Done"
  SubList = SubList & "      <ul>" & vbNewLine
  
  For ParagraphNumber = 1 To Selection.Paragraphs.Count
    'MsgBox Selection.Paragraphs.Item(ParagraphNumber)
    ListItem = Left(Selection.Paragraphs.Item(ParagraphNumber).Range.Text, Len(Selection.Paragraphs.Item(ParagraphNumber).Range.Text) - 1)
    SubList = SubList & "        <li>" & ListItem & "</li>" & vbNewLine
  Next ParagraphNumber
  
  SubList = SubList & "      </ul>" & vbNewLine
  SubList = SubList & "    </li> <!-- remember to delete other '</li>' -->" & vbNewLine
  SubList = SubList & vbNewLine
  'MsgBox SubList
  
  Selection = SubList
End Sub
Sub Make_Selected_UnorderedList()

  Dim UnorderedList As String
  
  UnorderedList = UnorderedList & "  <ul>" & vbNewLine
  UnorderedList = UnorderedList & Selection
  UnorderedList = UnorderedList & "  </ul>" & vbNewLine
  'MsgBox UnorderedList
  Selection = UnorderedList
End Sub
Sub Make_Selected_Section()

  Dim Section As String
  Dim SectionTitle As String
  
  SectionTitle = Selection.Paragraphs.Item(1).Range.Text
  SectionTitle = Left(SectionTitle, Len(SectionTitle) - 1)
  
  Section = Section & "    <section>" & vbNewLine
  Section = Section & "      <title>" & SectionTitle & "</title>" & vbNewLine
  Section = Section & vbNewLine
    
  Section = Section & Selection
  Section = Section & vbNewLine & "    </section>"
  
  Selection = Section
  Selection.Style = "Done"
  Selection.Paragraphs.Item(4).Range.Delete
  
End Sub
Sub Make_Selected_Paragraph()

    Dim Paragraph As String
    
    Paragraph = Left(Selection, Len(Selection) - 1)
    Paragraph = "    <p>" & Paragraph & "</p>"
    Paragraph = Paragraph & vbNewLine
    
    'MsgBox Paragraph
    Selection.Style = "Done"
    Selection = Paragraph
End Sub
