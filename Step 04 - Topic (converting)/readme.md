
The *main* problem with converting legacy Word documents into any other format, is that the underlying document (what you can't see or get access to) is often 'messed up'. For example, if you look at the ``(1) Overview`` document, you'll have no-doubt about it having two bulleted lists.

![image](overview.png width=50% height=50%)

However, if you run:

```
Sub Count_Lists()
  MsgBox ActiveDocument.Lists.Count, , "Count"
End Sub
```

It will report only one list:

![image](count.png)

