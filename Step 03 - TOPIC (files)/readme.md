*This application assumes that the content of each TOPIC is the content between each heading with the 'Heading 1' style.*

Steps
-----
1. Open `C Lesson 1.docx`
2. Run the `Main()` sub-procedure in `topics.bas` - it will:
   * Add a meta-data table to the end of the document showing:
     *  The text of each 'Heading 1'
     *  The paragraph number of each 'Heading 1'
     *  The paragraph number of the paragraph before each 'Heading 1'

![image](meta_data_table.png)

   * Opens a new document for each 'Heading 1', then:
     * Copies its content into the new document
     * Offers to save the new document with the text of the 'Heading 1' as the filename
     * Closes the new document
   * Repeats this process for each 'Heading 1' in `C Lesson 1.docx`

![image](exported_topic_files.png)
