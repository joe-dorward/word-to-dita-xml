Sub Test_Get_Topic_Header()
  MsgBox Get_Topic_Header("Topic", "Outline Topic (Topic)"), , "Test_Get_Topic_Header()"
End Sub
Function Get_Topic_Header(TopicType As String, TopicTitle As String) As String

  Dim TopicHeader As String
  Dim Topic_ID As String
  
  TopicHeader = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine
  TopicHeader = TopicHeader & "<!DOCTYPE "
  TopicHeader = TopicHeader & LCase(TopicType)
  TopicHeader = TopicHeader & " PUBLIC ""-//OASIS//DTD DITA "
  TopicHeader = TopicHeader & TopicType
  TopicHeader = TopicHeader & "//EN"" "
  TopicHeader = TopicHeader & """" & LCase(TopicType) & ".dtd" & """>" & vbNewLine
  
  ' add topic element
  TopicHeader = TopicHeader & "<" & LCase(TopicType)
  
  ' get topic id
  Topic_ID = TopicTitle
  Topic_ID = LCase(Topic_ID)
  Topic_ID = Replace(Topic_ID, " ", "_")
  Topic_ID = Replace(Topic_ID, "(", "")
  Topic_ID = Replace(Topic_ID, ")", "")
  
  TopicHeader = TopicHeader & " id=" & """t_" & Topic_ID & """>" & vbNewLine
  
  ' add title element
  TopicHeader = TopicHeader & "  " & "<title>" & TopicTitle & "</title>" & vbNewLine
  
  ' get 'body' element
  Select Case TopicType
    
    Case "Topic"
      TopicHeader = TopicHeader & "  " & "<body>" & vbNewLine
    
    ' add other cases for other topic-types
        
  End Select
  
  Get_Topic_Header = TopicHeader
End Function
