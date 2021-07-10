#  VBA of Wall Sreet

## Overview of Project
The objective of this project is to help Steve analyze the trend of twelve green energy company stocks for 2017 and 2018. Steve is helping his parents make a wise decision about investing money in purchasing stocks of a company named Daqo (ticker: DQ). We focused on obtaining the total daily volume and annual return of these twelve companies from raw data, using VBA (Visual Basic for Applications)

### Purpose
The purpose of this project is to learn coding in VBA as an introduction to programming. We learned different aspects of programming which includes: 
1. Writing macros ; 
2. Using _arrays_ ; 
3. Doing _loops_ and _nested loops_ ;
4. Use conditionals such as _if-then_ , _Elseif_ and _Else_; 
5. Making a good habit of formatting the programs so that others can follow the logic of the steps easily and the importance of "Comments". We were introduced to _pseudocode_ though we did not write it ourselves. 
6. For ease of viewing and using the program, we learned to make buttons to run the program, color code the cells based on conditions, make popup message box and making _input box_.
7. We learned to do Google search to find commands for certain tasks  if we are stuck. For example, we searched Google to find the command to obtain the end of the rows in a set of data from Stackoverflow.com.
8. Finally, we learned to obtain the running time of the program so that we can check the efficiency of the code written. 

## Analysis and Challenges
We were given an Excel file that we converted into a Macro enabled format and named it *VBA_Challenge.xlsm*. We imported *VBA_Challenge.vbs* in the macro of the file and modified it according to the directions given. The challenge was to refactor the script by using arrays so that the program would run much faster compared to not using arrays originally. The macro, _AllStocksAnalysisRefactored_ was partially filled. Our goal was to fill in the blank with arrays, initialize variables, write some _if_ conditions according to the instructions given in the challenge. 

My largest challenge in completing the **_Deliverable 1_** was understanding the concept of arrays and how it sores the data in bins. Once that was understood, the rest of the codes logic seemed straight forward. 


## Results
I created a new worksheet, **_Refactor Challenge_** for my output. ( Please note that I kept the practice worksheet, _All Stocks Analysis_ unchanged so that I could refer to it as I was learning to write the VBA code challenge). I was able to reproduce the example of the final Excel table given with and without arrays. The program ran more than five times faster when I used arrays to store data in bins. The results of the refactored time from the popup message box for the stocks in 2017 and 2018 shown below. These image files are included in the _Resources_ folder.


<p style="text-align: Left ;"> <strong> 2017 </strong>
<img src='./resources/VBA_Challenge_2017.png'>

<p style="text-align: Left ;"> <strong> 2018 </strong>
<img src='./resources/VBA_Challenge_2018.png'>


## Conclusion

The challenge was to organize and present the past stock market data to Steve so that he can help his parents make a wise decision on buying stocks. We showed the (1) Total Daily Volume and (2) Yearly Return of twelve green energy stocks for 2017 and 2018. We presented the data color coded so that Steve can glance at the outcome by color. We also installed buttons for ease of use. Finally, as a challenge for a programmer, we refactored the original code to make the program run faster to improve efficiency. 

## Advantage of refactoring code:

 By refactoring data in array, we sorted data in bins which allowed the computer to run faster. I can definitely see the advantage of such refactoring when dealing with a large data set. We only focused on the running time for this challenge. One could also refactor to analyze data differently. This is advantageous because it is not necessart to write a program from scratch. This program is short, therefore there are less issues. I can imagine that the a complicated and long program will not work the first time it is written. A refactoring will most likely be necessary.

## Disadvantage of refactoring code:

I had a hard time understanding the purpose and usage of various parts of the  refactoring instruction. This makes me think that refactoring could get complicated if the person refactoring is not able to find ways to make the whole code work together. The end result could be worse. This could be an issue if the program is large and have several steps of calculations. A long and complicated program will very likely have many steps that are depdent on each other. I think that refactoring may potentially mess up certain parts of the program if it is not done correctly.
