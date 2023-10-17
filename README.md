# 𝐌𝐗-𝟒𝟏𝟎𝟏_𝐆𝐫𝐨𝐮𝐩𝟓_𝐈𝐧𝐟𝐨𝐫𝐦𝐚𝐭𝐢𝐨𝐧-𝐃𝐚𝐭𝐞-𝐚𝐧𝐝-𝐓𝐢𝐦𝐞-𝐚𝐧𝐝-𝐋𝐨𝐨𝐤𝐮𝐩-𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬
𝑨𝒖𝒕𝒉𝒐𝒓/𝒔: 𝘑𝘰𝘩𝘯 𝘙𝘦𝘺 𝘋𝘦𝘤𝘢𝘯𝘰, 𝘚𝘩𝘦𝘳𝘪𝘭𝘺𝘯 𝘎𝘰𝘯𝘻𝘢𝘭𝘦𝘴, 𝘢𝘯𝘥 𝘍𝘳𝘪𝘵𝘻 𝘎𝘢𝘣𝘳𝘪𝘦𝘭 𝘗𝘢𝘭𝘮𝘢

Information, Date and time, and Lookup Functions are highlighted in this section along with information on how to utilize and manipulate them in Microsoft Excel. Every definition of a function includes a reference to its syntax.

#### 𝐀. 𝐈𝐍𝐅𝐎𝐑𝐌𝐀𝐓𝐈𝐎𝐍 𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬

  A.1. ISERROR

-> The Excel ISERROR function returns TRUE for any error type excel generates, including #N/A, #VALUE!, #REF!, #DIV/O!, #NUM!, #NAME?, or #NULL!

  𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘌𝘙𝘙𝘖𝘙 (𝘷𝘢𝘭𝘶𝘦)


  A.2. ISERR

-> The Excel ISERR function returns TRUE for any error type except the #N/A error.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘌𝘙𝘙(𝘷𝘢𝘭𝘶𝘦)


  A.3. ISNA

-> The Excel ISNA function returns TRUE when a cell contains the #N/A error and FALSE for any other value, or any other error type.

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘕𝘈(𝘷𝘢𝘭𝘶𝘦)


  A.4. ERROR.TYPE

->  The Excel ERROR. TYPE function returns a number that corresponds to a specific error value. You can use ERROR. TYPE to test specific kinds of errors.

->  If no error exists, ERROR. TYPE returns #N/A.

->  See left for a key to the error codes returned by ERROR.TYPE
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘌𝘙𝘙𝘖𝘙.𝘛𝘠𝘗𝘌(𝘷𝘢𝘭𝘶𝘦)
 

  A.5. ISNUMBER

->  The Excel ISNUMBER function returns TRUE when a cell contains a number, and FALSE if not.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘕𝘜𝘔𝘉𝘌𝘙 (𝘷𝘢𝘭𝘶𝘦)

  A.6. ISEVEN

-> The Excel ISEVEN function returns TRUE when a numeric value is even, and FALSE for odd numbers.

-> ISEVEN will return the #VALUE error when a value is not numeric.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =I𝘚𝘌𝘝𝘌𝘕(𝘷𝘢𝘭𝘶𝘦)

  A.7. ISODD

-> The Excel ISODD function returns TRUE when a numeric value is odd and FALSE for even numbers.

-> ISODD will return the #VALUE error when a value is not numeric.

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘖𝘋𝘋 (𝘷𝘢𝘭𝘶𝘦)

  A.8. ISBLANK

-> The Excel ISBLANK function returns TRUE when a cell contains is empty, and FALSE when a cell is not empty.

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘉𝘓𝘈𝘕𝘒(𝘷𝘢𝘭𝘶𝘦)

  A.9. ISLOGICAL

-> The Excel ISLOGICAL function returns TRUE when a cell contains the logical values TRUE or FALSE, and returns FALSE for cells that contain any other value, including empty cells.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘓𝘖𝘎𝘐𝘊𝘈𝘓(𝘷𝘢𝘭𝘶𝘦)

  A.10. ISTEXT

-> The Excel ISTEXT function returns TRUE when a cell contains a text, and FALSE if not. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘛𝘌𝘟𝘛(𝘷𝘢𝘭𝘶𝘦)

  A.11. ISNONTEXTV

-> The Excel ISNONTEXT function returns TRUE for a nontext value, for example, a number, a date, a time, etc. 

-> The ISNONTEXT function also returns TRUE for blank cells and for cells with formulas that return nontext results. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘕𝘖𝘕𝘛𝘌𝘟𝘛(𝘷𝘢𝘭𝘶𝘦)

  A.12. ISREF

-> The Excel ISREF function returns TRUE when a cell contains a reference or space, and FALSE if not. 

->  You can use the ISREF function to check is a cell contains a valid reference. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘙𝘌𝘍(𝘷𝘢𝘭𝘶𝘦)

  A.13. ISFORMULA

-> The Excel ISFORMULA function returns TRUE when a cell contains a formula, and FALSE if not. 

-> When a cell contains a formula ISFORMULA will return TRUE regardless of the formula’s output or error conditions. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘍𝘖𝘙𝘔𝘜𝘓𝘈(𝘷𝘢𝘭𝘶𝘦)

  A.14. TYPE

-> The Excel TYPE function returns a numeric code representing “type” in 5 categories:
Number = 1, Text = 2, Logical = 4, Error = 16, and Array = 64.

-> Use TYPE when the operation of a formula depends on the type of value in a particular cell. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘛𝘠𝘗𝘌(𝘷𝘢𝘭𝘶𝘦)


#### 𝐁. 𝐃𝐀𝐓𝐄 & 𝐓𝐈𝐌𝐄 𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬
  B.1. 

  B.2. 

  B.3. 

  B.4 

  B.5

  B.6
  
  B.7
  
  B.8
  
  B.9
  
  B.10
  
  B.11
  
  B.12
  
  B.13
  
  B.14
  
#### 𝐂. 𝐋𝐎𝐎𝐊𝐔𝐏 𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬
  C.1 

  C.2 

  C.3 

  C.4 

  C.5

  C.6
  
  C.7
  
  C.8
  
  C.9
  
  C.10
  
