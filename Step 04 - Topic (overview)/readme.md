
The *main* problem with converting legacy Word documents into any other format, is that the underlying document (what you can't see or get access to) is often 'messed up'. For example, if you look at the ``(1) Overview.docx`` document, you'll have no-doubt about it having two bulleted lists.

<img src="overview.png" width="25%" height="25%">

However, if you run:

```
Sub Count_Lists()
  MsgBox ActiveDocument.Lists.Count, , "Count"
End Sub
```

It will report only one list:

![image](count.png)

**Conclusions:** 

* It's unlikely that we can rely on simple VBA sub-procedures to automatically convert legacy Word documents into valid DITA XML files

* Some kind of semi-automatic, ad-hoc, interventionist, process will likely be neccessary - for example - selecting a few paragraphs and running a specific sub-procedure over them

### Solution

* Run the ``main()`` sub-procedure in ``topic_overview.bas``
* Select the first sub-list (and) run ``Make_Selected_SubList()`` 
* Select the second sub-list (and) run ``Make_Selected_SubList()``
* Select the first list (and) run ``Make_Selected_UnorderedList()``
* Select the second list (and) run ``Make_Selected_UnorderedList()``
* Select the paragraphs between the ``<body>`` and ``</body>`` elements (and) run ``Make_Selected_Section()``
* Select the "Important Notes:" paragraph (and) run ``Make_Selected_Paragraph()``

*That should produce an almost-correct DITA XML topic file, that needs little-more than a light-edit to make it correct.*  
