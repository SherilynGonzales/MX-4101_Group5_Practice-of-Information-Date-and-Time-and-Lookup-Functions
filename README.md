# 𝐌𝐗-𝟒𝟏𝟎𝟏_𝐆𝐫𝐨𝐮𝐩𝟓_𝐈𝐧𝐟𝐨𝐫𝐦𝐚𝐭𝐢𝐨𝐧-𝐃𝐚𝐭𝐞-𝐚𝐧𝐝-𝐓𝐢𝐦𝐞-𝐚𝐧𝐝-𝐋𝐨𝐨𝐤𝐮𝐩-𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬
𝑨𝒖𝒕𝒉𝒐𝒓/𝒔: 𝘑𝘰𝘩𝘯 𝘙𝘦𝘺 𝘋𝘦𝘤𝘢𝘯𝘰, 𝘚𝘩𝘦𝘳𝘪𝘭𝘺𝘯 𝘎𝘰𝘯𝘻𝘢𝘭𝘦𝘴, 𝘢𝘯𝘥 𝘍𝘳𝘪𝘵𝘻 𝘎𝘢𝘣𝘳𝘪𝘦𝘭 𝘗𝘢𝘭𝘮𝘢

Information, Date and time, and Lookup Functions are highlighted in this section along with information on how to utilize and manipulate them in Microsoft Excel. Every definition of a function includes a reference to its syntax.

### 𝐀. 𝐈𝐍𝐅𝐎𝐑𝐌𝐀𝐓𝐈𝐎𝐍 𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬

  A.1. ISERROR

> -> The Excel ISERROR function returns TRUE for any error type excel generates, including #N/A, #VALUE!, #REF!, #DIV/O!, #NUM!, #NAME?, or #NULL!

  𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘌𝘙𝘙𝘖𝘙 (𝘷𝘢𝘭𝘶𝘦) >
> ![ISERROR](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/ed794bfa-f91f-4e79-bd1e-e2234b6b973f)


  A.2. ISERR

> -> The Excel ISERR function returns TRUE for any error type except the #N/A error. 
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘌𝘙𝘙(𝘷𝘢𝘭𝘶𝘦)
> ![ISERR](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/4ffe6c4e-36d8-4fba-acb1-a8a814939f86)


  A.3. ISNA

> -> The Excel ISNA function returns TRUE when a cell contains the #N/A error and FALSE for any other value, or any other error type. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘕𝘈(𝘷𝘢𝘭𝘶𝘦)
> ![ISNA](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/41ba1d45-271e-475e-8890-a9af19dfda36)


  A.4. ERROR.TYPE

> ->  The Excel ERROR. TYPE function returns a number that corresponds to a specific error value. You can use ERROR. TYPE to test specific kinds of errors.
>
> ->  If no error exists, ERROR. TYPE returns #N/A.
>
> ->  See left for a key to the error codes returned by ERROR.TYPE 
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘌𝘙𝘙𝘖𝘙.𝘛𝘠𝘗𝘌(𝘷𝘢𝘭𝘶𝘦)
> ![ERROR TYPE](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/d22f81b5-32ac-4644-8ba7-911d89a29c6d)


  A.5. ISNUMBER

> ->  The Excel ISNUMBER function returns TRUE when a cell contains a number, and FALSE if not.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘕𝘜𝘔𝘉𝘌𝘙 (𝘷𝘢𝘭𝘶𝘦)
> ![ISNUMBER](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/c8c1a3c5-4423-47fc-9cff-a88d62749dcb)

  A.6. ISEVEN

> -> The Excel ISEVEN function returns TRUE when a numeric value is even, and FALSE for odd numbers.
>
> -> ISEVEN will return the #VALUE error when a value is not numeric.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =I𝘚𝘌𝘝𝘌𝘕(𝘷𝘢𝘭𝘶𝘦)
> ![ISEVEN](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/e8b1c12d-f072-4eb2-9b8c-4805d57229c4)

  A.7. ISODD

> -> The Excel ISODD function returns TRUE when a numeric value is odd and FALSE for even numbers.
>
> -> ISODD will return the #VALUE error when a value is not numeric.

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘖𝘋𝘋 (𝘷𝘢𝘭𝘶𝘦)
> ![ISODD](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/563db462-1b86-4970-940e-7cd720eb7265)

  A.8. ISBLANK

> -> The Excel ISBLANK function returns TRUE when a cell contains is empty, and FALSE when a cell is not empty.

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘉𝘓𝘈𝘕𝘒(𝘷𝘢𝘭𝘶𝘦)
> ![ISBLANK](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/20405a1f-8ec0-4305-94d8-7cc80f139a44)

  A.9. ISLOGICAL

> -> The Excel ISLOGICAL function returns TRUE when a cell contains the logical values TRUE or FALSE, and returns FALSE for cells that contain any other value, including empty cells.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘓𝘖𝘎𝘐𝘊𝘈𝘓(𝘷𝘢𝘭𝘶𝘦)
> ![ISLOGICAL](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/36712ed6-23d1-4bab-8d1c-9c2d0c705c84)

  A.10. ISTEXT

> -> The Excel ISTEXT function returns TRUE when a cell contains a text, and FALSE if not. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘛𝘌𝘟𝘛(𝘷𝘢𝘭𝘶𝘦)
> ![ISTEXT](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/1508d5fb-a693-41ee-a0d4-ede620bcf8ae)

  A.11. ISNONTEXT

> -> The Excel ISNONTEXT function returns TRUE for a nontext value, for example, a number, a date, a time, etc.
>
> -> The ISNONTEXT function also returns TRUE for blank cells and for cells with formulas that return nontext results. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘕𝘖𝘕𝘛𝘌𝘟𝘛(𝘷𝘢𝘭𝘶𝘦)
> ![ISNONTEXT](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/55a4a52d-4f9f-4d84-876f-533878a1383a)

  A.12. ISREF

> -> The Excel ISREF function returns TRUE when a cell contains a reference or space, and FALSE if not.
>
> ->  You can use the ISREF function to check if a cell contains a valid reference. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘙𝘌𝘍(𝘷𝘢𝘭𝘶𝘦)
> ![ISREF](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/d5730a19-94bd-4756-ae5c-567b0cf45b69)

  A.13. ISFORMULA

> -> The Excel ISFORMULA function returns TRUE when a cell contains a formula, and FALSE if not.
>
> -> When a cell contains a formula ISFORMULA will return TRUE regardless of the formula’s output or error conditions. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘐𝘚𝘍𝘖𝘙𝘔𝘜𝘓𝘈(𝘷𝘢𝘭𝘶𝘦)
> ![ISFORMULA](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/a98bdcee-1d9b-4508-8e03-c05c4d06871f)

  A.14. TYPE

> -> The Excel TYPE function returns a numeric code representing “type” in 5 categories:
Number = 1, Text = 2, Logical = 4, Error = 16, and Array = 64.
>
> -> Use TYPE when the operation of a formula depends on the type of value in a particular cell. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘛𝘠𝘗𝘌(𝘷𝘢𝘭𝘶𝘦)
> ![TYPE](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/d7dbaaf0-a90d-4124-9d63-fce3c2b69cc0)


### 𝐁. 𝐃𝐀𝐓𝐄 & 𝐓𝐈𝐌𝐄 𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬
  B.1. DATE

> ->  The Excel DATE function creates a valid date from the individual year, month, and day components.
>
> ->  The DATE function is useful for assembling dates that need to change dynamically based on other values in a worksheet

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘋𝘢𝘵𝘦(𝘺𝘦𝘢𝘳,𝘮𝘰𝘯𝘵𝘩,𝘥𝘢𝘺)
> ![DATE](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/52840872-9ae0-4308-babe-89036a4b1a61)

  B.2. TIME

> ->  The Excel TIME function is a built-in function that allows you to create a time with individual hour, minute, and second components.
>
> ->  The TIME function is useful when you want to assemble a proper time inside another formula.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    = 𝘛𝘪𝘮𝘦(𝘩𝘰𝘶𝘳,𝘮𝘪𝘯𝘶𝘵𝘦,𝘴𝘦𝘤𝘰𝘯𝘥)
> ![TIME](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/835dfbca-4845-47c7-830f-df2ac370f5e5)

  B.3. DateValue

> ->  The Excel DATEVALUE function converts text that appears in a recognized format (i.e. a number, date, or time format) into a numeric value.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘋𝘢𝘵𝘦𝘷𝘢𝘭𝘶𝘦("𝘺𝘦𝘢𝘳-𝘮𝘰𝘯𝘵𝘩-𝘥𝘢𝘺")
> ![DATEVALUE](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/3982b82a-4752-481b-aa7d-3abf85f1d8fb)

  B.4. TimeValue

> -> The Excel TIME VALUE function converts a time represented as text into a proper Excel time. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘛𝘐𝘔𝘌𝘝𝘈𝘓𝘜𝘌("𝘵𝘪𝘮𝘦_𝘵𝘦𝘹𝘵")
> ![TIMEVALUE](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/f27e88ac-ef69-4eee-bdf9-6382c17e29ef)

  B.5. Now&Today

> ->  The Excel NOW() function returns the current date and time, updated continuously when a worksheet is changed or opened.
>
> ->  The Excel TODAY() function returns the current date, updated continuously when a worksheet is changed or opened.
>
> Note: Both functions take no arguments.

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘕𝘖𝘞() 

    =𝘛𝘖𝘋𝘈𝘠()
> ![NOW TODAY](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/a6178a3e-dd2b-4fff-91ba-af58418f8eb7)

  B.6. Hour, Minute, Second

>> HOUR
>>
>> ->  The Excel HOUR function returns the hour component of a time as a number between 0 and 23. For example, with a time of 9:30 AM, HOUR will return  at 9,
>>
>> Serial Number
>>
>> ->  Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations.
>>
>> ->  By default, January 1, 1900, is serial number 1, and January 1, 2008, is serial number 39448 because it is 39,448 days after January 1, 1900.
>>

>> MINUTE
>>
>> -> The Excel MINUTE function extracts the minute component of a time as a number between 0 and 59.
>>
>> -> For example, with a time of 9:34 AM, a minute will return 34.
>> 

>> SECOND
>>
>> -> The Excel SECOND function returns the second component of a time as a number between 0 and 59.
>>
>> -> For example, with a time of 9:10:15 AM, the second will return 15. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘏𝘖𝘜𝘙(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳)

    =𝘔𝘐𝘕𝘜𝘛𝘌(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳)

    =𝘚𝘌𝘊𝘖𝘕𝘋(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳)
>> ![HOUR,MINUTE,SECOND](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/2c7e5214-b7b2-43d5-82d8-8dfc579d6f7e)

  B.7. Day, Month, Year

>> DAY
>>
>> -> The Excel DAY function returns the day of the month as a number between 1 to 31 from a given date.
>>
>> -> You can use the DAY function to extract a day number from a date into a cell.
>>

>> MONTH
>>
>> -> The Excel MONTH function extracts the month from a given date as a number  between 1 to 12.
>>

>> YEAR
>>
>> -> The Excel YEAR function returns the year component of a given date as a 4-digit number. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘋𝘈𝘠(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳)
  
    =𝘔𝘖𝘕𝘛𝘏(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳)

    =𝘠𝘌𝘈𝘙(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳)
>> ![DATE,MONTH,YEAR](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/6fd695ed-35cc-4644-b6d7-a47f9267e7f8)

  B.8. Weeknum

> -> The Excel WEEKNUM function takes a date and returns a week number (1 54) that corresponds to the week of the year.
>
> -> The WEEKNUM function starts counting with the week that contains January 1.
>
> -> By default, weeks begin on Sunday.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘞𝘌𝘌𝘒𝘕𝘜𝘔(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳)
> ![WEEKNUM](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/fe0f9936-c6c3-4421-9eda-6931e79ce24b)

  B.9. Weekday

> -> The Excel WEEKDAY function takes a date and returns a number between 1 and 7 representing the day of the week.
>
> -> By default, WEEKDAY returns 1 for Sunday and 7 for Saturday.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘞𝘌𝘌𝘒𝘋𝘈𝘠(𝘴𝘦𝘳𝘪𝘢𝘭_𝘯𝘶𝘮𝘣𝘦𝘳,[𝘳𝘦𝘵𝘶𝘳𝘯_𝘵𝘺𝘱𝘦])
> ![WEEKDAY](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/3d399807-6f75-4ea8-8f78-8dede6b654bd)

  B.10. EDATE

> -> The Excel EDATE function returns a date on the same day of the month, in months in the past or future.
>
> -> You can use EDATE to calculate expiration dates, maturity dates, and other due dates.

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘌𝘋𝘈𝘛𝘌(𝘴𝘵𝘢𝘳𝘵_𝘥𝘢𝘵𝘦,𝘮𝘰𝘯𝘵𝘩𝘴)
> ![EDATE](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/7f02718e-b4fc-4be3-a9b3-ea8830332d1b)

  B.11. EOMONTH

> -> The Excel EOMONTH function returns the last day of the month. 

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘌𝘖𝘔𝘖𝘕𝘛𝘏(𝘴𝘵𝘢𝘳𝘵_𝘥𝘢𝘵𝘦,𝘮𝘰𝘯𝘵𝘩𝘴)
> ![EOMONTH](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/c82ec7bb-941c-48c6-b432-d00cac29d53b)

  B.12. Workday

> -> The Excel WORKDAY function takes a date and returns the nearest working day in the future or past.
>
> -> You can use the WORKDAY function to calculate things like ship dates, delivery dates, and completion dates that need to take into account working and nonworking days.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘞𝘖𝘙𝘒𝘋𝘈𝘠(𝘴𝘵𝘢𝘳𝘵_𝘥𝘢𝘵𝘦, 𝘥𝘢𝘺𝘴, [𝘩𝘰𝘭𝘪𝘥𝘢𝘺𝘴])
> ![WORKDAY](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/189f6fe4-c727-4855-938d-7d25ff3ac1f6)

> Parameters
>
> -> Start date - The date from which to start.
>
> -> days - The working days before or after start_date
>
> -> holidays - [optional] A list of dates that should be considered non-working days.

  
  B.13. WORKDAY.INTL

> -> The Excel WORKDAY.INTL function takes a date and returns the nearest working in the future or past, based on an offset value you provide.
>
> -> Unlike the WORKDAY, WORKDAY.INTL allows you to customize which days are considered weekends (non-working days).
  
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘞𝘖𝘙𝘒𝘋𝘈𝘠.𝘐𝘕𝘛𝘓(𝘴𝘵𝘢𝘳𝘵_𝘥𝘢𝘵𝘦, 𝘥𝘢𝘺𝘴, [𝘸𝘦𝘦𝘬𝘦𝘯𝘥], [𝘩𝘰𝘭𝘪𝘥𝘢𝘺𝘴])
>  ![WORKDAY INTL](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/d10ded5a-1017-434c-8eb3-57f835fdac54)

> Parameters
>
> -> start_date The start date. ? days - The end date.
>
> -> weekend - [optional] Setting for which days of the week should be considered weekends.
>
> -> holidays - [optionall A list of one or more dates that should be considered non-working days.

  B.14. Days

> -> The Excel DAYS function returns the number of days between two dates.
>
> ->  With a start date in A1 and end date in B1, = DAYS(B1,A1) will return the days between the two dates.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘋𝘈𝘠𝘚(𝘦𝘯𝘥_𝘥𝘢𝘵𝘦, 𝘴𝘵𝘢𝘳𝘵_𝘥𝘢𝘵𝘦)
> ![DAYS](https://github.com/John-Rey-Decano/MX-4101_Group5_Practice-of-Information-Date-and-Time-and-Lookup-Functions/assets/143807174/744b1f6b-a58b-4e0b-a877-52591d45ad07)

  
### 𝐂. 𝐋𝐎𝐎𝐊𝐔𝐏 𝐅𝐮𝐧𝐜𝐭𝐢𝐨𝐧𝐬
-> The Excel LOOKUP function performs an approximate or exact match lookup in a one-column or one-row range, and returns the corresponding value from another one-column or one-row range.

-> LOOKUP’s default behavior makes it useful for solving certain problems in Excel. 

-> can be used to find the email addresses of 1000 employees from the contact list. Or can check specific details of an employee from a wide array of lists.

  C.1. LOOKUP 
> -> The Microsoft Excel LOOKUP function returns a value from a range (one row or one column) or from an array.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘓𝘖𝘖𝘒𝘜𝘗( 𝘷𝘢𝘭𝘶𝘦, 𝘭𝘰𝘰𝘬𝘶𝘱_𝘳𝘢𝘯𝘨𝘦, [𝘳𝘦𝘴𝘶𝘭𝘵_𝘳𝘢𝘯𝘨𝘦] )


  C.2. VLOOKUP
> -> VLOOKUP is an Excel function to look up data in a table organized vertically.
>
> -> It supports approximate and exact matching, and wildcards (* ?) for partial matches

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘝𝘓𝘖𝘖𝘒𝘜𝘗( 𝘷𝘢𝘭𝘶𝘦, 𝘵𝘢𝘣𝘭𝘦 , 𝘪𝘯𝘥𝘦𝘹 ,[𝘳𝘦𝘴𝘶𝘭𝘵_𝘳𝘢𝘯𝘨𝘦] )
 
Parameters
>
> -> value - The value to look for in the first column of a table.
>
> -> table - The table from which to retrieve a value.
>
> -> index - The column in the table from which to retrieve a value.
>
> -> result range - [optional] TRUE = approximate match (default). FALSE = exact match.


  C.3. HLOOKUP
> -> HLOOKUP is an Excel function to look up data in a table organized horizontally.
>
> -> It supports approximate and exact matching, and wildcards (* ?) for partial matches
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘏𝘓𝘖𝘖𝘒𝘜𝘗( 𝘷𝘢𝘭𝘶𝘦, 𝘵𝘢𝘣𝘭𝘦 , 𝘪𝘯𝘥𝘦𝘹 ,[𝘳𝘦𝘴𝘶𝘭𝘵_𝘳𝘢𝘯𝘨𝘦] )


  C.4. MATCH

> -> The Excel MATCH function returns the position of an item in a range
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘔𝘈𝘛𝘊𝘏(𝘭𝘰𝘰𝘬𝘶𝘱_𝘵𝘺𝘱𝘦, 𝘭𝘰𝘰𝘬𝘶𝘱_𝘢𝘳𝘳𝘢𝘺, 𝘮𝘢𝘵𝘤𝘩_𝘵𝘺𝘱𝘦


  C.5. CHOOSE

> -> The CHOOSE function in Excel is designed to return a value from the list based on a specified position.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    = 𝘊𝘏𝘖𝘖𝘚𝘌(𝘱𝘰𝘴𝘪𝘵𝘪𝘰𝘯, 𝘷𝘢𝘭𝘶𝘦1, [𝘷𝘢𝘭𝘶𝘦2, ... 𝘷𝘢𝘭𝘶𝘦_𝘯]

Parameters
>
> -> position - The position number in the list of values to return. It must be a number between 1 and 29.
>
> -> value1, value2, ... value, n - A list of up to 29 values. A value can be any one of the following: a number, a cell reference, a defined name, a formula/function, or a text value


  C.6. AREAS

> -> The AREAS function is a built-in function in Excel that is categorized as a Lookup/Reference Function

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘈𝘙𝘌𝘈𝘚(𝘳𝘦𝘧𝘦𝘳𝘦𝘯𝘤𝘦)

  
  C.7. ADDRESS

> -> The Excel ADDRESS function returns the address for a cell based on a given row and column number.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    = 𝘈𝘋𝘋𝘙𝘌𝘚𝘚( 𝘳𝘰𝘸, 𝘤𝘰𝘭𝘶𝘮𝘯, [𝘳𝘦𝘧_𝘵𝘺𝘱𝘦]. [𝘳𝘦𝘧_𝘴𝘵𝘺𝘭𝘦]. [𝘴𝘩𝘦𝘦𝘵_𝘯𝘢𝘮𝘦] )

Parameters
>
> -> row num - The row number to use in the cell address.
>
>-> col_ num - The column number to use in the cell address.
>
>-> ref_type - Optional. It is the type of reference to use. If this parameter is omitted, it assumes that the ref_ type is set to 1.
>
>-> ref_style - Optional. It is the reference style to use: either A1 or R1C1. If this parameter is omitted, it assumes that the ref__style is set to TRUE.
>
>-> sheet_name - Optional. It is the name of the sheet to use in the cell address. If this parameter is omitted, then no sheet name is used in the cell address.

  
  C.8. COLUMN

> -> The Excel COLUMN function returns the column number for reference.
 
𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘊𝘖𝘓𝘜𝘔𝘕 ([𝘳𝘦𝘧𝘦𝘳𝘦𝘯𝘤𝘦])

  
  C.9. COLUMNS

> -> The Excel COLUMNS function returns the count of columns in a given reference

𝑺𝒚𝒏𝒕𝒂𝒙

    =𝘊𝘖𝘓𝘜𝘔𝘕𝘚 (𝘢𝘳𝘳𝘢𝘺)

  
  C.10. INDIRECT

> -> The Excel INDIRECT function returns the reference to a cell based on its string representation.

𝑺𝒚𝒏𝒕𝒂𝒙

    = 𝘐𝘕𝘋𝘐𝘙𝘌𝘊𝘛( 𝘴𝘵𝘳𝘪𝘯𝘨_𝘳𝘦𝘧𝘦𝘳𝘦𝘯𝘤𝘦, [𝘳𝘦𝘧_𝘴𝘵𝘺𝘭𝘦] )

Parameters
>
> -> string reference - A textual representation of a cell reference.
>
>-> ref_style - Optional. It is the reference style to use: either A1 or R1C1. If this parameter is omitted, it assumes that the ref, style is set to TRUE.


  
