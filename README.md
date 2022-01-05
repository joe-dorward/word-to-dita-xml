## Converting Word documents to DITA XML using VBA sup-procedures


### STEP 01

1. Open `Word To DITA XML.docx`
2. Run the `Main()` sub-procedure in `bookmap.bas` - it will:
   * Open a new Word document
   * Write BOOKMAP boiler-plate into that Word document
   * Add chapter-elements from the headings in the Word document with the 'dita-chapter' style
   * Then offer to save the Word document as `b_word_to_dita_xml.docx` - if you save the Word document with the `.txt` extension, you can change it later to `.ditamap`.

----
