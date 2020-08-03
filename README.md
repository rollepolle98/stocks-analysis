# stocks-analysis

## Overview of the Project

In this project the task was to initially create a workbook for the client, Steve. Steve was tasked with helping his parents analyze different types of green energy stock. The information that we start with is that Steve's parents invested in DACO New Energy Corporation (DQ). We are tasked with helping Steve look into DQ by comparing the stock data to other similar companies. Initially we started with finding the Total Daily Volume (TDV) and Yearly Return Percentage (YRP) for DQ in order to analyze the stocks trading frequency along with value from the beginning to end of the year. After Further analyzation, it became clear that DQ may not be the best option, so we started doing the same for all other green energy stocks we were provided with. Once we accumulated all of this information, we were able to create a workbook that included a flexible macro which allowed Steve look at data sets for different years. From there we went on to creating a more user-friendly experience by adding static and conditional formatting to easier display the stocks TDV and YRP, along with adding two different run buttons which allowed Steve to analyze stock for each year at the click of a mouse.

Now that we created a workbook for Steve to analyze the current data at hand, we then went onto refactoring the script in order to make it more applicable for future and larger data sets. The main purpose of refactoring the script is to make the code more efficient by taking fewer steps, using less memory, or improving the logic of the code which can be visually verfied by examing whether or not it successfully made the VBA script faster than it did initially. Looking below at the visuals, we can further analyze whether we were successful or not.


## Results

### VBA Script Speed - 2017 - Original

![](Resources/VBA_Challenge_2017(Original).png) 

##### Elapsed Run Time = 3.878906 Seconds

### VBA Script Speed - 2018 - Original

![](Resources/VBA_Challenge_2018(Original).png) 

##### Elapsed Run Time = 4.785156 Seconds

### VBA Script Speed - 2017 - Refactored

![](Resources/VBA_Challenge_2017(Refactored).png) 

##### Elapsed Run Time = 0.109375 Seconds

### VBA Script Speed - 2018 - Refactored

![](Resources/VBA_Challenge_2018(Refactored)_2018.png) 

##### Elapsed Run Time = 0.1054688 Seconds

### Looking at the images above, we can clearly see that the refactored code for 2017 was 3.769531 seconds faster and 4.6796872 seconds faster for 2018. Further analyzation shows the refactored script run about 35 times faster for 2017 and 45 times faster for 2018. On average the refactored code runs about 40 times faster than the original.  

### Based on the visuals and analysis, it’s clear that the refactored code helped increase the speed of the script making it more efficient and applicable to larger data sets.

## Summary

### 1) What are the advantages or disadvantages of refactoring code?

The advantages to refactoring code is that it improves the design, efficiency, and application on larger scale. This is important because it helps makes software easier to understand which in turn helps finding bugs therefore allowing program to run faster.

The only disadvantage I can think of is that it just adds a lot more time to the process so say if one is on a tight budget or tight schedule, refactoring may cause some issues short term but in the long run, I still believe it’s worth it.

### 2) How do these pros and cons apply to refactoring the original VBA script?

Well to start off, the advantages listed above all apply to refactoring the original VBA. Refactoring the script allowed the speed of the script to increase on average 40 times faster, which helps it be more efficient. The script is much easier to understand when look at the code which is always an advantage. Lastly refactoring not only improved the speed and design but also the application in that the script can be applied to larger data sets without losing large amount of efficiency.

The cons also apply in that refactoring the code added a significant amount of time to my analysis but in the end I believe that it is ultimately worth it.
