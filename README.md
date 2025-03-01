# Microsoft-Excel-Project-6

# Case study

A colleague, Lucas, has asked you to update a spreadsheet called Reseller Details that records details of Adventure Work’s resellers in the United States. This information in the spreadsheet was downloaded from another system. The download process created several inconsistencies or errors within the data.

These errors include unnecessary spaces, the use of the wrong case, and entries that need to be joined together or split apart. 
______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________
# Standardizing Text-Based Data in Excel

# Overview: 
In this exercise, I applied text functions in Excel to clean and standardize reseller information. The tasks included removing unnecessary spaces, changing text case, extracting portions of data, and combining entries.

# Key Tasks Completed:
# 1. Remove Redundant Spaces:

      ⦁ Used the TRIM function to remove spaces before and after text while keeping necessary spaces between words.
      ⦁ Example Formula: =TRIM(B2)
      
# 2. Change Case of Entries:

      ⦁ Used the PROPER function to change the text case to capitalize the first letter of each word.
      ⦁ Example Formula: =PROPER(D2)
    
# 3. Extract Portions of Text:

      ⦁ Used LEFT to extract the first 6 characters from a string.
      ⦁ Example Formula: =LEFT(H2, 6)
      ⦁ Used RIGHT to extract the last 8 characters.
      ⦁ Example Formula: =RIGHT(H2, 8)
      ⦁ Used MID to extract characters from a specific position in the string.
      ⦁ Example Formula: =MID(H2, 8, 3)

# Combine Entries:

      ⦁ Used CONCAT to combine data from two cells with a space in between.
      ⦁ Example Formula: =CONCAT(G2, " ", I2)

# Transform Text Case:

      ⦁ Used UPPER to convert extracted text into uppercase.
      ⦁ Example Formula: =UPPER(L2)

# Final Steps:

      ⦁ Autofill: Applied the Autofill shortcut to quickly copy formulas down the columns.
      ⦁ Finalizing Data: Copied the results of formulas as values and deleted original columns with incorrect data.

# Conclusion:
By using a range of text functions, I successfully cleaned and standardized the data in the reseller worksheet.
