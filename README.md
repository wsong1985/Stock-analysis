# Stock-Analysis

## **Overview of Project**

### To assist Steve in analyzing the dataset of the entire stock market over the last few years by utilizing a refactored code based on the original solution code.


## **Results**

### **Stock Performance between 2017 and 2018**

  _1. Upon comparing the stock performance, we can see that all the stocks had a negative return in 2018, except for ENPH and RUN._

  _2. ENPH continued to retain a positive return rate in 2018, which was relatively lower than in 2017._

  _3. RUN outperformed the other stocks in 2018, with an approximately 78.5% increase in return compared with 2017._

  _4. TERP continued to underperform in 2018, with a 5% loss in return; however, the loss was relatively minor compared to 2017._ 
  
  <table>
  <tr>
    <td>Stock Performance in 2017</td>
    <td>Stock Performance in 2018</td>
  </tr>
  <tr>
    <td><img src="Resources/Stock Performance _2017.png" width=300></td>
    <td><img src="Resources/Stock Performance _2018.png" width=300></td>
  </tr>
 </table>


### **Execution Time Comparison between the Original Script and the Refactored Script**

  _1. For the year 2017, the execution time of the original script was about 0.796875 seconds. With the refactored script, the execution time can be shortened to 0.1289063 seconds, approximately six times faster than the original script._ 
  
  <table>
  <tr>
    <td>Original Script</td>
    <td>Refactored Script</td>
  </tr>
  <tr>
    <td><img src="Resources/Elapsed Run Time with Original Code _2017.PNG" width=300></td>
    <td><img src="Resources/VBA_Challenge_2017.PNG" width=300></td>
  </tr>
  </table>

  _2. For the year 2018, the execution time of the original script was about 0.765625 seconds. With the refactored script, the execution time can be shortened to 0.1289063 seconds, approximately six times faster than the original script._ 

  <table>
  <tr>
    <td>Original Script</td>
    <td>Refactored Script</td>
  </tr>
  <tr>
    <td><img src="Resources/Elapsed Run Time with Original Code _2018.PNG" width=300></td>
    <td><img src="Resources/VBA_Challenge_2018.PNG" width=300></td>
  </tr>
  </table>


## **Summary**

- **What are the advantages and disadvantages of refactoring code in general?**

  * Advantages of Refactoring Code:

    * _It can improve the design of the software._

    * _It simplifies the code and makes it easier to understand._

    * _It helps the developer find bugs._

    * _It enhances efficiency and helps the program run faster._ 
  
  * Disadvantages of Refactoring Code:

    * _Refactoring code is very time-consuming._

    * _There is a chance of creating bugs/errors during the refactoring process._

    * _It is risky when refactoring code during testing phase, which will lead to a situation that the developer has no clue on how to fix the new bugs/errors._

- **What are the advantages and disadvantages of the original and refactored VBA script?**

  * Original VBA script:
  
    * Advantanges:
    
      * _It has more flexibility, which is easier to add new functions to the existing code._
      
      * _Each function in the script acts independently, which is easier to modify the existing code._
      
    * Disadvantages:
    
      * _The code is potentially getting long._
      
      * _Some steps are redundant._
      
      * _Sometime, it is hard to follow._
    
  * Refactored VBA script:
  
    * Advantages:
    
      * _It runs faster than the original VBA script._
      
      * _It consists of fewer lines._
      
      * _It is easier to understand._
      
    * Disadvantages:
    
      * _It takes time to refactor the code._
      
      * _It requires some thinking, and it may need to redesign the existing functions so they can interact with each other more efficiently._ 
