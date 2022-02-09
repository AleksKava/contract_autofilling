## contract_autofilling
# Importing excel data to word template with python

This program was created to simplify the filling of standard agreements.
The user fills in a table in an Excel file “Data”. At the same time, he has the opportunity to choose from two different executors. After running the CAF.exe file, the program automatically generates an agreement with the necessary information entered into it, and also adjusts the content of the agreement depending on the executor selected by the user.
The finished contract is saved in a folder specially created for it. Together with the contract, the program creates a new excel file “Contract data”, which is then used by another program when forming expert conclusions on the basis of the agreement created by this program.
While creating the program, the following functions and modules taken from other sources were used:
- to write the number in words
https://github.com/seriyps/ru_number_to_text
- to delete paragraphs
https://stackoverflow.com/questions/29283306/deleting-paragraph-from-cell-in-python-docx
- to delete rows in the table
https://stackoverflow.com/questions/55545494/in-python-docx-how-do-i-delete-a-table-row
- to adjust excel columns widths
https://stackoverflow.com/questions/63493743/attributeerror-worksheet-object-has-no-attribute-set-column
