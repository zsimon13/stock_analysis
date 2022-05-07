# Green Stocks Analysis

## Project Overview
Steve was looking into stocks for his parents to invest in, and asked for my assitance in determining if it was worth it. In order to determine if the investment was worth while, I used the VBA in Microsoft excel to analyze the performance of 12 stocks. Sotcks performance was gauged by the stock's: 
- Daily volume  
- Annual return

### Purpose
The purpose of this project was to effeciently analyze multiple stocks using VBA. While my initial code was sufficient to analyze the 12 stocks in question; it became clear that there was room for improvement in code performance, and would *need* to be improved should we scale up and analyze more stocks. In order to maintain efficiency with a larger data set and increase efficiency with the current data set, I would need to refactor the code. This project will determine if refactoring the code achieved this goal.

## Results
### Refactoring the Code
In order to increase the efficiency of the analysis, I had to optimize how the loops were nested. This was achieved by first creating an array for the ticker symbols "ticker" and then declaring a variable "tickerindex", that could be used in the 3 arrays I would then create: tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
### Refactored Code
![Refactored Code](https://user-images.githubusercontent.com/102814578/167235068-629d4347-02ee-45ed-b022-0e808144c495.png)
![Refactored Code 2](https://user-images.githubusercontent.com/102814578/167235079-07a228fa-5507-49d6-8e74-6d336025e002.png)
### Original Code
![OG Code](https://user-images.githubusercontent.com/102814578/167235085-5ebcf555-b305-4d7a-8ac9-d21e25097ed1.png)
Using the "Tickerindex" variable allowed the code to assign the values of tickerVolumes, tickerStartingPrices, and tickerEndingPrice to the individual ticker symbols basically simultaneouly as it ran through the data set. The nested loop in the original code would have to run through the data set in order to read the values before it assigned them to the proper cells. This would in theory make the analysis run faster
###Run Time for Analysis
#### Original Code Run time
![2017 Stock Analysis Run Time](https://user-images.githubusercontent.com/102814578/167235369-443678c6-813e-4417-abc6-19db7c02baba.png)
![2018 Stock Analysis Run Time](https://user-images.githubusercontent.com/102814578/167235374-8514f321-459b-459a-904d-cf1cde649b3f.png)
#### Refactored Code Run Time
![VBA_Challenge_2017](https://user-images.githubusercontent.com/102814578/167235389-679b5247-0e9f-4a35-bc19-fe10c57dc1dc.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/102814578/167235390-464e3a5e-02fe-4713-99d3-bce3853d5bc6.png)
As you can see, the run time was almost 10 times as fast after refactoring the code.

