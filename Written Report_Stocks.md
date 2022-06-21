
# Stock VBA Challenge

## Overview of Project

### Purpose

     - In order to satisfy the Steve's request to analyze the entire stock market over the last few years, which includes thousands of stocks, it became necessary to refactor our code to loop through the dataset in a more efficient manner. Without refactoring, it could take the code too long and the VBA script wouldn't run quickly. We will be refactoring our code to see if our stock analysis runs faster.
     - Prior to this, the primary information we collected and calculated for Steve were percent return and Total Daily Volume. To do so, we looped through all of the stocks, calculated the yearly returns of DQ's stock, learned that Daqo would potentially be a bad financial investment, and then analyzed multiple stocks to better advise which companies to invest in.
     - We calculated total daily volume and return and edited our analysis so Steve would easily be able to analyze differing years. However, in our initial presentation, our VBA code ran in about 0.5 seconds. By refactoring the code and making it more efficient, we will hopefully be able to present Steve with an even faster running code.

## Results

-  2017
      -  2017's dataset shows that the clean energy industry had a great year of return.
  
    ![screenshot3](https://user-images.githubusercontent.com/102992388/174891639-e94b0902-20b1-4963-832f-541d9d1db962.png)

- 2018
  
    - After a successful year, the industry had a bad year in 2018. Only Sunrun Inc and Enphase Energy Inc had a positive year in terms of return.
  
    ![screenshot4](https://user-images.githubusercontent.com/102992388/174891776-f7ff338d-8117-4495-92b2-3ac3a31a65c4.png)

- Run Time Comparison
  - The dataset from 2018 ran the macros much quicker than the dataset from 2017. 
    ![screenshot5](https://user-images.githubusercontent.com/102992388/174893378-8fe44508-98a4-4109-b0f1-5edb8bede590.png)
    ![screenshot4](https://user-images.githubusercontent.com/102992388/174893234-8660732c-ff82-4656-87c9-217cd05346e0.png)


## Analysis


     - In order to analyze the entire stock market, we created a tickerIndex variable and created three arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. I created an Input Box with the message: "What year would you like to run the analysis on?" so Steve can compare differing years.
     - I wanted to loop through the tickers and then loop through rows in the dataset
     - I created a for loop to initialize the tickerVolumes to zero for the 12 tickers we had. Next, I wanted this analysis to loop through all the rows in the dataset and increase volume for the current ticker with the code: 

``` 
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, "H").Value
```  
     
     - This allowed me to increase volume for each ticker. This code allowed me to track the total volume for each ticker. I ended up using Cells(i,"H").Value because I found it easier to track the columns by their assigned letter, rather than counting the alphabet (i.e. Cells(i,8).Value).
     - Afterwards, in order to ensure that we were properly counting the right ticker, I used the following code to keep track of the selected ticker:
 
   ![screenshot](https://user-images.githubusercontent.com/102992388/174883523-c643d2d0-3869-4b20-9859-bf6c033c7d3f.png)

     - This code checks to see if the current row is the first row, and then the last row with the selected ticker. If it doesn't match, we increased the tickerIndex.
     - Afterwards, I looped through the arrays to output Ticker, Total Daily Volume, and Return.
  
   ![screenshot2](https://user-images.githubusercontent.com/102992388/174884105-3625e034-f44e-4dc5-853f-9f5d94004d8f.png)

    
### Challenges 

    - I had the biggest trouble with steps 3b, 3c, and 3d. It took a lot of trial and error, but in the end I learned a lot about how to use conditionals to check if a current row matches with a row before and after it to then add to the selected tickerIndex.
  
### Questions
    - What are the advantages and disadvantges of refactoring code in general?
      - For both 2017 and 2018, my refactored code ran faster. Initially, my code ran at around 0.5 seconds to run through the macros. With the refactored code, it ran in 0.06 and 0.22 seconds.
      - A good advantage was also being able to review the concepts to ensure I understood what we were doing in the module correctly. Without the extra challenge of refactoring code, I would have had a more superficial understanding.
      - One disadvantage of refactoring code that only took 0.5 seconds to begin with would be the extra time that it took for us to refactor this code. Is the difference between 0.5 and 0.22 that noticeably different? However, I imagine with larger sets of code, this would become more noticeable.
    - What are the advantges and disadvantages of the original and refactored VBA script?
      - There aren't that many advantges and disadvantages that come to mind other than the ones I outlined above. However, one thing to note is in the refactored code, I liked that the button included all of the different macros we had learned up until that point. 
