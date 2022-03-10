# contract_autofilling
## Importing excel data to word template using python and python-docx

This program was created to simplify the filling of standard agreements.

The user fills in a table in an excel file “Data”. At the same time, he has the opportunity to choose from two different executors. After running the CAF.py file, the program automatically generates an agreement with the necessary information entered into it, and also adjusts the content of the agreement depending on the executor selected by the user.

Together with the agreement, the program creates a new excel file “Contract data”, which is then used by another program when forming expert conclusions on the basis of the agreement created by this program. Both files, the word "Contract" and the excel “Contract data” are saved in a folder specially created for it.

The user has also an option to select or cancel the sound notification at the end of the program execution.

Sound effects files are located in the folder System\Sounds.

The contract template is located in the folder System\Template.

While creating the program, the following functions and modules taken from other sources were used:
- to write the number in words
https://github.com/seriyps/ru_number_to_text
- to delete paragraphs
https://stackoverflow.com/questions/29283306/deleting-paragraph-from-cell-in-python-docx
- to delete rows in the table
https://stackoverflow.com/questions/55545494/in-python-docx-how-do-i-delete-a-table-row
- to adjust excel columns widths
https://stackoverflow.com/questions/63493743/attributeerror-worksheet-object-has-no-attribute-set-column
