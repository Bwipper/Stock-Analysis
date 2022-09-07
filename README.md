# The Analysis of Stocks from 2018
### Creating a macro in VBA to analyze stocks

## Overview of Project
  ### Purpose and Backround
    This project had us working in Excel and VBA to create a macro that our client can utilize to analyze stocks. This will allow them to make informed decisions about where to invest in the future. 

## Anaylsis and Challenges 
  ### Overview of Anaylsis 
    We used stock data from 2017 and 2018 to make sure our macro was correctly able to analyze the stocks trends of the past. This would confirm that we could take the macro and apply it to future stock data and inform users where they should and shouldnt be investing their money. 
    
  ### Challenges involved
    I am very new to the world of creating macros and writing code so this was quite a challenge for me. Thankfully the modules are pretty precise in their instructions so I could refer back to them and the previous work we did when I ran into issues. I also have a friend who works in SQL at a large bank who was in town and I was able to get some of her help on this project as well. She had never used VBA before but I guess if you learn one coding language the others are pretty similar. Ultimatly I think I just need to put the hours in and keep trying to learn this code and the languages going forward. I think with enough time I will get the hang of it and become profecient. I think the main issue is that I am not really thinking "in code", and thus I tend to overanalyze what Im doing ang get stuck. 
  
## Results
The image below shows the final result of our macro and its findings. You can see that many stock are currently in the negative but two stand out as growing rapidly. This is what our macro was intended to do to assist in making informed decisions around investing money in stocks. 

<img width="1576" alt="VBA_Challenge_Stocks Image" src="https://user-images.githubusercontent.com/111392120/188963421-87b9e914-2f91-41bb-bede-6263e49b6e1a.png">

We later took this code and added a timer to see how long it takes to run and analyze the stocks it is referencing. In the image below you can see the origianl time it took to run this code. 

<img width="457" alt="VBA_Challenge_Not _Refactored" src="https://user-images.githubusercontent.com/111392120/188963824-afd2daa4-be02-4c1a-8e80-983d51171720.png">

We wanted to try to make this code as efficent as possible so it is able to analyze much larger data sets and not take a long time to do it. We went back and refactored the code, simplifying parts of it and adding more efficent code where it was needed. The result of this new code is that it became almost twice as fast as the original. See the image below and compare it to the one above. 

<img width="436" alt="VBA_Challenge_Refactored" src="https://user-images.githubusercontent.com/111392120/188964691-b3412b9a-99de-4bd2-ac65-6a41527e1be5.png">

It doesnt matter too much in situations like this where there isnt a massive amount of information and it only takes seconds to run. But when working with very large data sets the difference in run times would be very noticible. 
