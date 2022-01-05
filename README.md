## Converting Word documents to DITA XML using VBA sup-procedures


### STEP 01

1. Open `Word To DITA XML.docx`
2. Run the `Main()` sub-procedure in `bookmap.bas` - it will:
   * Open a new Word document 
   * Write BOOKMAP boiler-plate into it
   * Add chapter-elements to it (corresponding to the headings in `Word To DITA XML.docx` with the 'dita-chapter' style
   * Then offer to save it as `b_word_to_dita_xml.docx` - if you save the Word document with the `.txt` extension, you can change it later to `.ditamap`.

----
