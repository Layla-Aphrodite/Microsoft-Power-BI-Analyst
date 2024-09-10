# Date calculations

### Introduction

Entering the current date and time is one of the more common tasks that you’ll complete in your worksheets. These entries are also valuable resources in a Microsoft Excel spreadsheet. For example, if there are date or time entries in the data, you can perform time-based analysis like showing sales totals broken down by month or year. Date and time entries are also important because they are stored as numbers in Excel and can be used in calculations. 

Excel contains a range of functions designed to work on calendar or clock entries. You can explore the full list of these in the date and time category in the function library. This reading demonstrates a selection of functions that you might find useful in your day-to-day spreadsheets, alongside a brief overview of how date and time are tracked in Excel.

### How Date and Time are tracked 

Excel stores dates and times as serial numbers. Excel took January 1st, 1900 as the first date and assigned it serial number 1. Excel interacts with the system clock in the PC to keep track of time changes. It generates a new serial number when it perceives a change. The recognized serial number for the current date increases by one every twenty-four hours, and Excel treats time as a decimal fraction of the full 24-hour day.

When you are entering a date in an Excel worksheet, it should be typed as numbers with the month, day, and year entries separated by forward slashes. An example of this could be 08/30/23. You can change the appearance of the date with formatting. 

Time entries are also typed as numbers with the hours and minutes separated by a colon, such as 10:30. Time can be formatted to display as a 12-hour clock or a 24-hour clock.

Date and time entries typed directly into the spreadsheet are not automatically dynamic. Like any other number, they remain as they were originally entered. However, other date-generating functions will dynamically update using the system clock as a reference.

### TODAY and NOW

TODAY and NOW are both examples of functions that generate dynamic results.

A cell that contains the TODAY formula always shows the current system date and increments every twenty-four hours.

**=TODAY()**

A cell that contains the NOW formula always shows the current system time and increments whenever the formulas are recalculated in the spreadsheet.

**=NOW()**

The syntax for these two functions is similar because they do not require any arguments. However, they still require parentheses after their respective names. There should not be any characters between the parentheses, not even a space.

The default setting for the result of the NOW function formula is to display both the date and the time. This can be customized using the Time number formatting choices.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/aiAeCN4IQ4SnufgD1JybmA_625c73fc92434e3a84473df8539771e1_image.png?expiry=1726099200000&hmac=mt83DBBtZaWPF90dmMw2o8breOCy0nAqR-blS4eV4gk'>

### DAY, MONTH and YEAR

When analyzing dates, it’s useful to be able to manipulate data based on a portion of a date. For example, you could manipulate data to find all sales completed in 2023 or all purchases made in the month of August. 

You can extract date data using the DAY, MONTH and YEAR functions. These functions extract the relevant component of a date and display it in a separate cell. You can then use this date element for sorting or filtering.

In the Excel screenshot below, the date has been entered correctly in cell A2 as numbers separated by a forward slash. The date is in MM/DD/YY format.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/vt7RRRZsRDyjtuDO6bcLKA_40adf6eed74f40868a7638666f97fce1_image.png?expiry=1726099200000&hmac=ow0U7Wdgk7Fif3ZajT5uOZwqlQ_4JeQO7MvS-9Y3LBw'>

As shown in the above screenshot, the following formula, which uses DAY, will extract and display the day component of the full date value in cell A2:

**=DAY(A2)**

This next formula, which uses MONTH, will extract and display the month component of the full date in cell A2:

**=MONTH(A2)**

And this third formula, which uses YEAR, will extract and display the year component of the full date in cell A2:

**=YEAR(A2)**

### NETWORKDAYS

Excel spreadsheets are an effective way to plan a project. For example, if you need to know the number of days available to work on a project, then it’s useful to be able to calculate a result in working days rather than calendar days.

NETWORKDAYS is a function that calculates the number of working days between a start date and an end date. It does not include weekend days in the result it generates.

NETWORKDAYS uses the country settings on your PC to determine which days constitute a weekend. In the Adventure Works example below, the country is set to be USA. So, the weekend days to be excluded are Saturday and Sunday.

What if you need to specify a different day, or days, for the weekend? In these instances, you can use a different function called NETWORKDAYS.INTL. This version of the function makes no assumptions about the weekend. Instead, you must include an argument in the formula which instructs Excel what days to exclude.

Neither NETWORKDAYS nor NETWORKDAYS.INTL allows for statutory public holidays as there is no information on these days built into the functions. Instead, when creating a formula using these functions, you must include a table of public holidays and reference it in the formula.

In the screenshot example below, the formula in cell C2 takes the entry in A2 as the start date and the entry in B2 as the end date.



<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/t3wR0imhRDuL3rnPjEO5Eg_3233051878644d6d8b01a4c59a6d5ae1_image.png?expiry=1726099200000&hmac=oM_ZU-039oh1ueh4l0dUruQyE5P7sATFmqp7cAu2Ny4'>

As the screenshot above shows, the formula calculates that there are 170 days until the end of the year. When calculating this result, it ignored any date that fell on a Saturday or Sunday. The formula in cell E2 takes the same start and end date but adds a third argument which asks Excel to also exclude any relevant holiday dates in the cell range H2 to H12. Excel ignores any dates in the past and then excludes the future dates in cells H5 to H12, giving an adjusted result of 162 days. 

If the start date in A2 had been created with a TODAY function formula, then the number of working days left to the end date would update every twenty-four hours.

### DATE
You have already used the CONCAT function to combine multiple text entries into one single cell entry. The DATE function allows you to perform a similar operation on date entries. If the month, day and year numbers are in separate cells, the DATE function can be used to combine them. The advantage of this function is that it generates a result which is also a date.

In this example, the month, day and year have been entered into cells A2, B2 and C2. The three arguments in the DATE function formula tell Excel where the Year, Month and Day entries are. The result is then a date.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/1Zzv7HOIT4C1fdj_gs7HgA_314ce60e771143ee93abbc49966dffe1_Date-and-Time-calc-image2.png?expiry=1726099200000&hmac=y4ysXZnz9SpbldUK1VAOvoqEf5pXZLnDprnnZi0A4DA'>

### DATEDIF
The DATEDIF function calculates the number of days, months or years between two dates. It is an older function referred to as a “legacy” function in Excel. It is still supported, but you might notice that Excel does not provide a floating help message to help you create it. Similarly, if you are using the Insert function wizard, you won’t see this function listed. However, it is still a common function in spreadsheets that track information such as duration of employment. If you intend to use this function, there are some situations where results are not reliable. Please refer to the Microsoft support page DATEDIF function referenced in additional resources.

In this example, the DATEDIF calculation in D2 takes the entry in B2 as a start date and the entry in C2 as an end date. The date in cell C2 is generated with the TODAY function, so it’s dynamic. The third argument in the formula, “y”, tells Excel to display the difference between the start and end dates in whole years.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/u4BYNthaRfa6_03-55wYmQ_ba9cb33680814919b841cec6c97dcee1_image.png?expiry=1726099200000&hmac=qEsSVEaoB5_TxNqYKDbiDCITowqv3KFC89D6DsPgq9A'>
