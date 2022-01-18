*This application assumes that the content of each TOPIC is the content between each heading with the 'Heading 1' style.*

Steps
-----
1. Open `C Lesson 1.docx`
2. Run the `Main()` sub-procedure in `topics.bas` - it will:

(https://github.com/joe-dorward/word-to-dita-xml/blob/main/Step%2002%20-%20TOPIC%20(files)/meta_data_table.png)

   * Add a meta-data table (see `meta_data_table.png`) to the end of the document showing:
     *  The text of each 'Heading 1'
     *  The paragraph number of each 'Heading 1'
     *  The paragraph number of the paragraph before each 'Heading 1'
   * Opens a new document for each 'Heading 1', then:
     * Copies its content into the new document
     * Offers to save the new document with the text of the 'Heading 1' as the filename
     * Closes the new document
   * Repeats this process for each 'Heading 1' in `C Lesson 1.docx` (see `exported_topic_files.png`)
[meta_data_table.png]
