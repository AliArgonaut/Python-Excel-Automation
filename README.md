CLICK MAIN.PY TO SEE THE CODE 

SCENARIO: you are given a raw CSV file every day with customer information, the money they spent today, and the total money theyve spent before today. 
your goal is to create an easier to read excel report with their full name, email, customer ID, and an updated total amount theyve spent, and email this file to johnsmithers@gmail.com (a fake email I made)

this python script takes the CSV file, inputs that raw data into excel. From there it:

1. creates a new sheet
2. populates with relevant headers
3. combines the first and last name fields of the raw data and puts it in a full name field in the new sheet
4. copies over their emails
5. uses XLOOKUP to find and populate their IDs 
6. calculates the sum of all theyve spent and puts it into the total spent field
7. formats the entire new sheet to look prettier
8. emails the entire thing to john smithers with a short description and subject line describing what was done 

the benefit of this sample script is its reusability. all you need to do is take any different CSV with the same layout, put it next to the program, run the executable program, and the same operations are performed and the thing is emailed in literal seconds. What would take maybe 30+ minutes to do is done instantly and with no room for human error at all.

In a hypothetical where I am exposed to a routine where I know what data I'm working with and we're looking at multiple excel sheets and complex operations, we can imagine building out these programs to do more and more, especially if we use them in conjunction with ExcelScripts as well. 
