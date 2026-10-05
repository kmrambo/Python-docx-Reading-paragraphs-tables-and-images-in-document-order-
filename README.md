# Python-docx (Reading paragraphs,tables and images in document order)

## Description
python-docx is a Python library for creating and manipulating Microsoft Word (.docx) files.
The Python-docx package cannot read paragraphs, tables and images altogther in document order. In document order, either it can only render all the paragraphs at once or all tables at once or all images at once. Here, I provide a way in which paragraphs, tables and images present in a docx file can be read in document order into a dataframe in python.

The code is present inside the file named "Para_table_image_extraction.py". While running this code with any docx file as input, this code genertes 3 dataframes namely ```combined_df```, ```table_list``` ( a list basically) and ```image_df```.

All the contents of the docx file(including paragraphs, tables and images) are stored in a python dataframe called ```combined_df```. 
If an image is present after a paragraph in the document, a reference to the image will be stored in the ```combined_df``` dataframe, but not the actual image. You will have to refer the image index from ```combined_df``` and retrieve the image data from a separate dataframe called ```image_df``` which stores the image_index and the corresponding base64 encoded image data of each and every image that is present in the document.



Similarly, if a table is encountered in the document, the ```table_id``` column in ```combined_df``` dataframe gets filled up. And you will have to retrieve the corresponding table that is relevant to the ```table_id``` from ```table_list```.

I have illustrated this code on the following word document(docx):
![docx document](/images/snaphot_of_docx_document.png)

Below is a sample of ```combined_df``` dataframe that is created based on the above docx file.

![combined_df](/images/combined_df.png)

In the above dataframe:
1. ```para_text``` column denotes the actual paragraph content of the document. (Every row represents each paragraph in the document)
2. ```table_id``` column represents the table id of a table if a table is present in that location of the document. If no table is present, then its value is represented as "Novalue"
3. ```style``` represents the paragraph style of the corresponding paragraph

You should note that images and tables are not stored as such inside combined_df dataframe. A reference to those objects are only stored in combined_df. Which means for every image id or table id present in combined_df, you will have to retrieve either the image id or table id from ```combined_df``` and refer those ids from ```image_df``` for image data and ```table_list``` for table data.

Image files are represented in the following notation:
"**Document_Imagefile/image1.png/rId7/0** in ```combined_df```
which denotes that the reference is actually an image file which has "**image1.png**" as image name and "**rId7**" as the unique id or identifier for the image and "**0**" for image index to be referred in ```image_df```.

![image_df](/images/image_df.png)


Table objects are represented as "**<docx.table.Table object at 0x1020f1160>**" which denotes that a table is present at that location with the corresponding table_id. You can refer to this table_id in ```table_list```'s index to fetch the relevant table as a dataframe. ```table_list``` contains a list of all tables stored in dataframe format.

![table_list](/images/table_list.png)


I have not included the code to extract the images and tables from ```image_df``` and ```table_list``` using their unique identifiers from ````combined_df````. The reason I have not done this is because the use case might vary from case to case. You can extract them in your own way and do whatever successive operation you wish to. 

## How to use