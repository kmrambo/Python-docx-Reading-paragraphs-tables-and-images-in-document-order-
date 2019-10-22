# Python-docx (Reading paragraphs,tables and images in document order

python-docx is a Python library for creating and updating Microsoft Word (.docx) files.
The Python-docx package cannot read paragraphs, tables and images in document order. It can only render all the paragraphs at once or all tables at once or all images at once. Here, I provide a way in which paragraphs, tables and images present in a docx file can be read in document order into a dataframe in python.

All the contents of the docx file(including paragraphs, tables and images) are stored in a python dataframe called combined_df. If an image is present after a paragraph in the document, the index of the image will be stored in the combined_df dataframe, but not the actual image. You will have to refer this image index from combined_df and retrieve the image data from a separate dataframe called image_df which stores the image_index and the corresponding base64 encoded image data of each and every image that is present in the document.




Similarly, if a table is encountered in the document, the table_id column in combined_df dataframe gets filled up. 
