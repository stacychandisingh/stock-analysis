# stock-analysis
# An Analysis of Green Energy Stocks
Performing analysis on stock data to uncover the best investment options

# VBA Challenge

### Overview of Project
The objective of this project is to explore green energy stock performance by analyzing financial data using VBA.

### Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

When comparing stock performance between the years 2017 and 2018, it was observed that stocks performed better in 2017 than in 2018. In 2018, only two stocks performed positively. While in In 2017, only one stock of the 12 stocks dropped in performance.
 
The execution time of the original script proved to be shorter in time when compared to the refactored script; the refactored script took a longer time to execute. Typically, a refactored script generally has a shorter execution time. The following images show the short and long execution times of the original script and the refactored script for 2018, respectively:

The Original Script Output:

![original_script](resources/Original Script_2018.png)


The Refactored Script Output:

![VBA_Challenge_2018](resources/VBA_Challenge_2018.png)

Images for Deliverable 1 (2018 Output):

![VBA_Challenge_2018](resources/VBA_Challenge_2018.png)

Images for Deliverable 1 (2017 Output):

![2017](resources/VBA_Challenge_2017.png)

### Summary
In a summary statement, address the following questions:
#### What are the advantages or disadvantages of refactoring code?
	
The advantages of refactoring code is that it may be easier to understand, read, and maintain. Refactored code may also take a shorter time to execute upon running the code; the program runs faster. Disadvantages of refactoring code may include that it is time consuming to edit resulting in a longer process and there is a chance mistakes are made when refactoring the code and even more time is spent on resolving this new issue.


#### How do these pros and cons apply to refactoring the original VBA script?

The pros which apply may to refactoring the original VBA script include a faster code execution time upon running the refactored script and it may be easier to read with adequate comments included. The cons which apply to this refactoring exercise using the original VBA script is that a longer time was taken on finding ways of making the code more efficient by using the correct keywords; there is a chance that a programmer may lose the intended purposes of the code while refactoring. For this event, the refactoring exercise made the program's execution time longer than the original script execution time!
	

	A specific challenge I encountered was trying to remove the colored conditional formatting to the *Total Daily Volume* column even when the code seems clear as seen in the subsequent image: ![If_Statement](resources/If_Statement_Conditional Loop.png)
  
  The following image shows a portion of the refactored code which generated the correct total daily volume and percentage returns by setting the tickerIndex to zero before looping over the rows; the image also shows the creation of arrays by initializing the required variables by using the keyword, **Dim**:
  
  ![code](resources/Code Image.png)
  
  
  
  

 

