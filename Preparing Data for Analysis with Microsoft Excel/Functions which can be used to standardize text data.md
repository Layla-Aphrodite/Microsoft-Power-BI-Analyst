## The LEFT, RIGHT and MID Functions
The **LEFT, RIGHT and MID** functions are used to return a specific number of characters from either the left, the right, or the middle side of a cell entry. Typically, these functions are used in situations where you need to transfer parts of the cell content to a different column. 

The **LEFT** and **RIGHT** function formulas need two arguments. You must specify the cell which contains the original entry and the number of characters you would like extracted from it.

A **MID** function formula needs three arguments:

 - The first is the location of the text entry. 

 - The second is the position in the entry where Excel must begin counting characters. 

 - The third is the number of characters to the right of that position to extract.

Let’s explore an example of this using Adventure Works. Adventure Works is processing an order for Doctor Martin Garcia. While preparing this order, they must analyze and filter his name or title. This analysis would be easier if the data were formatted or standardized over several columns such as Title, First Name, and Last Name. 

In the image of the customer spreadsheet, the customer’s name Doctor Martin Garcia was typed in cell B2. If the spreadsheet were filtered by the term Doctor in the Title column, then this customer’s information would be included. If the term Doctor was part of a longer entry, then this customer might be overlooked in a filter or a search. 

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/tfPTs5f4RjO4G8eP4_s3dw_acc7c02904f14fc2a16e802ec5d5b6e1_image.png?expiry=1725926400000&hmac=KpWCFjBEl85o-XSoCGDUa62SP_oACIkYJX971rRBfxA'>


Worksheet with three-word entry in cell B2.

Let's use functions to standardize this text data.

The formula in C2 will extract the six characters to the left of the entry, the title Doctor. This formula reads as follows:

 **=LEFT(B2,6)** 

The formula in D2 will begin with the eighth character and return the next six characters to the right, which is the first name of Martin. This formula is:

**=MID(B2,8,6)**

The formula in E2 will extract the six characters to the right of the entry, which is the last name of Garcia. Excel counts the six characters from right to left in the cell. This formula is:

**=RIGHT(B2,6)**

Examples of Left, Mid and Right formula syntax displayed in cells.
Many data analysts use the LEFT, RIGHT and MID functions to split the contents of a column into three separate columns. With very large blocks of data, this may become clumsy.


<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/b53NhSanS9ynNVR-194H9A_a5108cda884e4c9bb6f626ccecc8e5e1_image.png?expiry=1725926400000&hmac=QUMXDywqMTu9tjHViEXWJI4tY5mMjjKn6Zc9XX7dYQM'>


There is an Excel feature called Text to columns which performs this action more efficiently. This tool is explained in the Microsoft support page 
Split text into different columns with the Convert Text to Columns Wizard
. It is also available in the Additional Resources reading at the end of this module.

## TRIM
The TRIM function is used to remove any empty spaces from text strings, except for the spaces between words. This is useful for situations where you suspect that there are random spaces at the beginning or end of an entry. If the TRIM function is run on a cell without any spaces before or after the entry, it will simply return the entry itself without making any changes. The TRIM function formula only needs one argument, which is the location of the piece of text you would like the function to work with. 

In the Image below, there are spaces before and after the entry Doctor Martin Garcia in cell B2. The formula in C2 will remove them but keep the spaces between the title, first, and last names. This formula reads as follows:

**=TRIM(B2)**

The formula result is the entry of Doctor Martin Garcia without the leading or trailing spaces. An example of this is displayed in D2 in the image below.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/b-V-hIp1T9mB1MooX4EGSA_04b26a296694448ba94cc69c6f25e7e1_image.png?expiry=1725926400000&hmac=bYjDeyQuGrcpTPV3ZkFTBvezQxPbei8FyasjqDHeL3s'>

Example of TRIM formula showing syntax.
This function is a useful way to tidy up a column of text before beginning any analysis. 

For example, you suspect that entries from A2 down to A200 in a worksheet contain unnecessary spaces. To fix this, you can use the TRIM function.

First you could Insert a new blank column to the right of the column which contains the inconsistent entries.  In this new column you could create a TRIM formula in the first cell  to remove any unnecessary spaces at the beginning and the end of the entry in A2. Excel will also remove extra spaces between words keeping only a single space between the words in the entry. Then you can use Autofill to copy the formula down to cell A200 to run a similar check and tidy-up process on the rest of the entries in column A.  You would then proceed to copy the results to the clipboard and then paste them back as values. And to finish off the TRIM procedure, you would then delete the original column A.

## UPPER, LOWER and PROPER
Using the wrong case in text data can make a summary or report appear untidy or unprofessional. There are three functions you can use to standardize the case used in text entries. These are UPPER, LOWER and PROPER.

The UPPER function converts all lowercase characters into uppercase, while the LOWER function does the opposite. The PROPER function will only capitalize the first character of each word in a piece of text. All three functions require only one argument, which is the location of the piece of text you would like the function to work on. 

In the following example, Doctor Martin Garcia’s name and title have been typed incorrectly in several ways in cells B2 to B4. The formula in C2 converts all the letters to uppercase. The formula is:

**=UPPER(B2)**

The sample result DOCTOR MARTIN GARCIA in D2 in the screenshot below displays the result this formula produces.

The formula in C3 using the LOWER function converts all the letters in the entry to lowercase. The formula is:

**=LOWER(B3)**

The result, doctor martin garcia, is displayed as a sample in D3 in the screenshot below.

The formula in C4 uses the PROPER function to create a result of three words with a capital at the beginning of each one. The formula reads as follows:

**=PROPER(B4)**

Cell D4 in the screenshot below displays the result of this formula: Doctor Martin Garcia.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/p50kegdFR8GDcmsG2QOLXg_3d517cd6f229437ca1ff7748776f93e1_image.png?expiry=1725926400000&hmac=ci1bnSMal46gFgIEPTEJCqTjckfOpM8cHW_IcBH0oQc'>

Examples of Upper, Lower and Proper formula syntax.


## CONCAT
There may be occasions when you need to combine entries from different cells in a spreadsheet into a single-cell entry. For this action, you can use the CONCAT function. The arguments for the function are the cell references separated by commas. As you type this formula, the floating Help box displays three dots to indicate that you can include multiple cell references.

<img src='https://d3c33hcgiwev3.cloudfront.net/imageAssetProxy.v1/jt2MAlxrSPObGZGEyK9XXA_90ee4ec23c8846bc81627a1087d3cbe1_image.png?expiry=1725926400000&hmac=H0vj-P7BVgAMeJtgbBDMLNklwCJCObxLufiBQJmerbg'>

CONCAT formula in progress with help message displayed.
You can use CONCAT to combine the content of cells that contain numbers. However, be aware that the use of the CONCAT function means that the data type of the combined result will be text and not numeric. 

Let’s take an example where cell A2 contains the title Doctor, B2 the first name Martin and C2 the last name Garcia. The formula in D2 reads as follows:

**=CONCAT(A2,B2,C2)**

A sample result of this formula is displayed in E2 in the screenshot below: DoctorMartinGarcia.

As you can imagine, this is not quite correct. Spaces are required between the words, but if you type spaces in the formula, it will generate an error. Instead, they must be added as arguments. They are surrounded by double quotes so that Excel knows that they are text characters to be added. The formula in D4 combines the same cells but also allows for spaces. The formula is:

=CONCAT(A2,” “,B2,” “,C2)

The result of this formula would be Doctor Martin Garcia, as is shown in the sample result in E4 in the screenshot below.

Two examples of CONCAT formula, one including spaces in result.
There is an older function in Excel called CONCATENATE which performs the same action as CONCAT. You may see this in spreadsheets. It is explained on the CONCATENATE function 
page
 on the Microsoft support site. It is also available in the Additional Resources reading.


