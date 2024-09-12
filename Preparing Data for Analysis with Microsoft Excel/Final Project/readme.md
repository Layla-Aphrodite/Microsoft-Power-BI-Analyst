# Creating an executive data summary

### Introduction

Over the past three weeks, you’ve learned a lot about Microsoft Excel. You know how to prepare worksheet data for presentation to a wider audience by using formatting or functions. You have also practiced creating a broad range of calculations.

In this exercise, you’ll draw on this knowledge to modify a worksheet and create a data summary for an executive audience. You will also create calculations to show the profit margin and performance compared to the previous year.

### Case study
Jamie, at Adventure Works, is attending a management meeting. She has been asked to prepare an Excel worksheet that presents sales figures for the first quarter of the year and compares these figures to the results for the same period in the previous year. 

This worksheet is called Summary and is in the workbook Quarter One Report.xlsx. In this worksheet, you’ll need to complete the following actions:

Create formulas that show the total quarter-one sales for both 2022 and 2023.

Create formulas that show the percentage increase in sales in 2023. 

And break down these totals by month with the use of further calculations.

Let’s help Jamie to complete this worksheet.

#### Step 1: Download the file.
Download and open the Microsoft Excel workbook Quarter One Report.xlsx. The workbook contains only one worksheet called Summary.

The sheet contains sales information for specific products spread over two years. It includes wholesale and retail prices as well as the quantity sold for each product. In the steps that follow, you need to reformat this information so that the worksheet shows only the required data and displays it effectively.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/kLrLv8OMQ3aRsvlfYOfDIg_64378990009d4e158a34bad04b16d2e1_image.png?expiry=1726272000000&hmac=7-PGuyMj9K-dPTI06bhiSNPdac1UYKCSenKm3D0BBAA'>

#### Step 2: Add and format headings.
1. Widen column A to accommodate the entries in cells A12 to A14. Add a new blank column to the left of column E. Column E is titled Product ID.

Tip: You can drag the right vertical line between the letters A and B manually, or you can use the shortcut double-click method to resize the column quickly.

2. In cell A4 type the heading TOTAL Q1 SALES and in A10 type the heading Q1 MONTHLY TOTALS.

3. Adjust the format of the heading in A4 as outlined below:

 - Increase the font size by one step.

 - Bold the text.

 - Add a background color.

 - Select the cells A4 to D4 and apply the Merge & Center option.

 - Next, use the Format painter option to apply the same formatting to cell A10.

Tip: Remember to position the cursor on the cell which already contains the formatting you want to copy before selecting Format painter.

4. Bold the headings in cells B5, C5 and D5 and turn on Wrap text. Use the Format painter to apply the same formatting to B11, C11, and D11.

#### Step 3: Customize and reorganize how the data is displayed. 
1. In cell H2, create a formula using PROPER to copy the product names in column G. Autofill the formula. Copy the results and use the Paste Values choice to replace the formulas with the result. Once this is done, delete column G.

Tip: The Paste values choice is available on both the Home ribbon and the right-click shortcut menu.

2. Highlight the block of cells F2 to Y246 and sort the block by Order Date. The sort order should be Oldest to Newest.

Tip: The headings and entries in columns A to E should not be sorted. Remember, you are highlighting a block of data. So, use the Sort dialog rather than the quick shortcut buttons to sort the data.

3. Hide columns F and columns S to Y.

4. Position the cursor on G2 and use the Freeze option on the View ribbon to freeze the columns to the left of the cursor and the row above the cursor.

#### Step 4: Use formulas to create new row information
1. Create a formula in K2 using MONTH and a formula in L2 using YEAR to extract the two component parts of the date in J2. Use Autofill to copy both formulas down to row 246.

Tip: You can highlight cells K2 and L2 and then double-click to Autofill them both at the same time.

2. In P2, create a standard multiplication formula that multiplies the retail price by the order quantity. Copy the formula using Autofill.

3. In cell Q2, create a formula using an IF function that calculates if tax is due on the amount in P2. 
 
 The IF function must check if the amount in P2 is over 2000. If it is, then the amount in P2 must be multiplied by 5%. If it’s not, then cell Q2 should display a 0. Use Autofill to copy this formula down.

Tip: The "Value if true" or "Value if false" entries for an IF can both be formulas if required. Remember that these formulas do not need an equals sign at the beginning. The standard rules for creating and controlling formulas will apply.

#### Step 5: Create formulas to calculate and compare the profit margin across two years.
1. In cell B6, use SUMIF to sum the sales values for 2022. The sales values are in the range R2 to R246. The criteria range will be L2 to L246. Create a similar formula in cell C6 but change the criteria to 2023.

Tip: Remember that in the SUMIF arguments, the criteria range comes first, then the criteria, and then the sum range.

2. In cell B12 use SUMIF to sum the range R2 to R103 if there is the number 1 in the criteria range K2 to K103. Add dollar signs to the R and K cell references so that the formula can be copied down with the cell references staying constant.

Tip: If you are using a standard keyboard, don’t forget that you can quickly add the dollar signs to a cell reference by placing the cursor on it and then pressing the F4 key on the keyboard.

3. Copy the formula from B12 into B13 and B14. In the B13 copy, change the criteria to 2. In the B14 copy, change the criteria to 3.

Tip: When you change the criteria in B13 to 2, a green triangle appears temporarily. This triangle is there to notify you that the formula in this cell is temporarily inconsistent with the formulas above and below. You do not need to act on this notification. It will disappear when you change the criteria in B14 to 3.

4. In C12 use SUMIF to sum the range R104 to R246 if it says 1 in the range K104 to K246. Add dollar signs to the R and K cell references.

5. Copy the formula to C13 and C14. In the C13 formula, change the criteria to 2, and in the C14 formula, change the criteria to 3.

6. Create a Percentage difference formula in D6 which shows the percentage by which sales increased in 2023. 

Tip: If you are unsure about how to create this calculation, you can check the syntax in the 
Useful Percentage Calculations
 
 reading. As a quick reminder, the logic is (New value-old value)/Old value.

7. Create a similar formula in D12. Copy the calculation in D12 down to D14.

### Conclusion
In this exercise, you made use of the wide range of techniques that you learned over the past few weeks. You have altered the formatting of the worksheet and modified its appearance to prepare it for presentation. You have used formulas, including logical and percentage formulas, to create an executive summary for the Quarter One performance for these products.
