# Replicating Excel Tables in Power BI with Selectable Columns

Recently I was tasked with converting some Excel tables into Power BI with the additional functionality of making the columns selectable. The Power BI extract would need to look exactly the same, this meant including the columns in the same order and keeping any formatting of the individual fields. There wasn't anything that covered the exact scenario so I decided to write this blog in hopes that it would help someone else.

I started with a dataset in a matrix that replicated the table structure of Excel (shown below). Each column had the individual field formatted to show the correct data format. The Product was displayed under the relevant product group and associated month.

![image](https://user-images.githubusercontent.com/126701906/228939744-b7b20c5d-1204-4e10-bbee-a5d45205c3bc.png)

The first step was to enable users to select the columns they wanted to be displayed in the table. I wanted to preserve the original format of the file in a separate input for additional analysis, so I started by creating a duplicate of the current table in Power Query Editor.

![image](https://user-images.githubusercontent.com/126701906/228939797-36c229c3-0ad0-409b-845a-f958ac8792cc.png)

The duplicate table was then unpivoted to show a list of values for column selection. In this case Purchased, Sold, Sale Total, Cost Total and Profit were unpivoted. 
Select the columns and right click the columns to unpivot:

![image](https://user-images.githubusercontent.com/126701906/228939839-0a57e89b-b063-4678-8c83-dba282890133.png)

This created an Attribute and Value column with all the field names from the unpivoted fields contained within the Attribute column:

![image](https://user-images.githubusercontent.com/126701906/228939887-0895584c-3f99-41f1-a2e2-2d444481fc7b.png)

After applying the changes in Power Query Editor, I replaced the fields in the matrix with the fields from the new table. Month, Product Group and Product were still in the original structure, so they just replaced those existing fields with the new duplicate table fields of the same name under the Row section. Then to replicate the same layout as the original the Attribute field was added to the Columns section and Value field to the Values section:

![image](https://user-images.githubusercontent.com/126701906/228939942-e50bc36d-26dc-4873-b393-2501eebebcb1.png)

By also including the Attribute field in a slicer the columns within the matrix become selectable, allowing you to activate or deactivate as many as you need:

![image](https://user-images.githubusercontent.com/126701906/228939987-431b43c8-80c1-4d66-bc28-32d071478bcd.png)

Looking at the full table with all columns selected (shown below), the columns are now in a different order and it has also lost the formatting for the individual fields. From Power Query Editor I could sort by ascending or descending on the Attribute column but this wouldn’t work for this particular table layout.

![image](https://user-images.githubusercontent.com/126701906/228940035-5ccaebf2-6ccd-4be6-93bc-09c3e8a2c7f1.png)

To amend the column order I created a new table in Power Query Editor using the selectable column names from the table in a Name field and added a sorting order ID field:

![image](https://user-images.githubusercontent.com/126701906/228940078-78e6f4fa-3997-4166-9b7a-35605f0b5609.png)

Then in the modelling view I linked the Name field from my sort order table to the Attribute field used in the unpivoted table:

![image](https://user-images.githubusercontent.com/126701906/228940111-cd0cf243-2d9f-4305-8c3a-88412ff729a1.png)

Switching back to the report view I created a calculated column for the sort order in the table with the Attribute column:

![image](https://user-images.githubusercontent.com/126701906/228940160-9dd89ec5-ff82-4195-842a-e4f3c8b653b4.png)

Selecting the Attribute column and using the Sort by Column function on the Column tools ribbon I was able to sort Attributes by the calculated sort order column:

![image](https://user-images.githubusercontent.com/126701906/228940196-1f4c4a76-ea40-4d89-823b-b3a0f238edc5.png)

This reordered the columns into the original order:

![image](https://user-images.githubusercontent.com/126701906/228940228-ca240d3c-0d0e-4543-ba7a-9a302f491bf5.png)

The final issue was to address the formatting that had been lost by converting multiple fields into one column.

For this to work I created a measure that looked for a string in the Attribute field and then formatted the value in the Value field. To be able to calculate totals at the month and product group level I used SUM within the FORMAT expression. 

In this table example I had two quantity fields which needed to be whole numbers and three currency fields (Profit, Total Cost and Total Sales). The DAX Measure below searches for the word ‘Profit’ or ‘Total’ within the Attribute and applies a format of currency to the Value field while retaining the ability to calculate the totals, to all other fields it applies a whole number while again retaining the ability to calculate totals.

DAX Expression:

    Value Measure For Multiple Data Formats = IF(OR(CONTAINSSTRING(SELECTEDVALUE(Table'[Attribute]),"Profit"),CONTAINSSTRING(SELECTEDVALUE('Table'[Attribute]),"Total")),FORMAT(SUM(Table'[Value])),"£#,##0.#0"),FORMAT(SUM('Table'[Value])),"#"))
  
Finally replacing the Value field in the Values Section of the matrix with the newly created measure returned the original layout now with the increased functionality of selectable  columns.

Final Output:

![image](https://user-images.githubusercontent.com/126701906/228940488-91976cb7-68ed-4fc1-9017-74a2f329b24e.png)
  
