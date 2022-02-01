# Stock Analysis with VBA

## Project Overview
Provide facilities to Steve (our customer) handling large stock of data in behalf of its research on analyzing stocks. Emphasize the importance of refactoring code and application of the VBA tools in data analysis to improve the efficiency of response. Analyze the impact to refactor a code in saving time and simplify the coding lecture and execution.
### Purpose
Make efficient the stock analysis by refactoring the code. The main purpose is to edit our initial VBA code and measure the performance. Compare both running times (initial VBA vs refactor VBA) to determine the advantages or disadvantages of refactoring code. Then establish the pros and cons that apply refactoring our initial VBA code. Finally, according to the feedback of response improvement with refactoring, analyze the stock market in 2017 and 2018 and the results of our data analysis.
## Analysis and Challenges
### Analysis of Refactoring Code

Here are the ways we work on our codes:

- **Without refactoring**
![VBA_Challenge_Unrefactored](https://user-images.githubusercontent.com/96077418/151905040-4c29e015-b35c-4e9b-b8a9-b80062d82129.png)

- **Refactored**
![VBA_Challenge](https://user-images.githubusercontent.com/96077418/151905047-d80ac043-3bc7-4f13-9162-678128f120cf.png)

If we see both pictures above, we can see that the unrefactored one has a code that is not simple to read. It shows steps that looks more complex even if they are very simple (talking about functionality). And by the other hand, the second image it is easier to read, we have a good structure on the code, simple steps and details on each function we want to do. If I am a new coder I can have a better approach on the second image than the first one, I can work better in the refactored code because I can understand easyly the goal of that script. I don't have to spend to much time in finding what is the purpose of hte code, like in the first image.

In addition when we run both codes, the program works faster with the second one than the first one. That is because it is also simple for the program to read the functions of the code and the program also doesn't needs to repeat functions like in the firs unrefactored code. That give us an efficient code and with a best response in behavior of the data analysis we need to do. The refactored code even works faster than the first one even we are handling large amounts of informations.

It is important to mention the time of response of each code (which we ll se lines below) but also the details on the refactored code. The details in green that describe the goal of each function. Why is so important? Because that bring us an architechture of what the coder is doing and how we can improve the work.

### Stock performance between 2017 and 2018

Once the unrefactored and the refactored codes where analyzed and compare, we can proceed to our data analysis.
The results that Steve will get using or cleaner code.
If we have a ook on the pictures below, we can see the information that Steve will get.
With three buttons he can decide in which stocks his parents could invest without spend time, days, months looking a large file of information.
They can do it in seconds.

So, Steve can see which stocks are better to invest on but also which year has best performance.
We can see that 2017 was a better year than the 2018, so they can make a bigger analysis whith more information on th stocks for 2017.
The returns we show in 2017 are so much better and also more stocks with possitive returns than the 2018.

![VBA_Challenge_2017_Outputs](https://user-images.githubusercontent.com/96077418/151905008-1e96122c-7e17-4e0f-b444-a505ad72346f.png)
![VBA_Challenge_2018_Outputs](https://user-images.githubusercontent.com/96077418/151905022-2f51fbf2-d1fb-44c3-821b-49d94fb45aff.png)

### Challenges and Difficulties Encountered

The big and one challenge we face was the time spent on refactoring the code. You need time to work on the code, look over on how to clean the code to improve the time of response. But it is very important that the time you spend working on the code adds value and not rests it. Once you get a cleaner code, the performance of it will be the best and easy to handle. If you don't refactor the code, you can face some technical issues or and overwork of the machine, and that will be worst because now you wont get the results on time for doing your anaylsys.

## Results
### Execution times original and refactored VBA script

Here we show the response times we mentioned before:

- **Unrefactored response time 2017**
![VBA_Challenge_2017 Unrefactored](https://user-images.githubusercontent.com/96077418/151065539-2fb89c7f-d0db-4d1e-8030-e1cee4a04ae1.png)

- **Refactored response time 2017**
![VBA_Challenge_2017](https://user-images.githubusercontent.com/96077418/151900692-b51e7d73-2c6c-46fb-98df-5a756fd2010f.png)

- **Unrefactored response time 2018**
![VBA_Challenge_2018 Unrefactored](https://user-images.githubusercontent.com/96077418/151900708-490a3908-6d7c-4e9b-b003-e2ba8c7c64a5.png)

- **Refactored response time 2018**
![VBA_Challenge_2018](https://user-images.githubusercontent.com/96077418/151900718-4cc0c80d-dff2-4f78-a86e-0a625302fe09.png)

It is very clear that the refactored code it is more efficient than the unrefactored one.
The times cannot be to representative because the amount of information is almost the same. But even that, it is clear the performance on both codes.
Nevertheless I am sure that if we increase the information data base with more stocks, we will see how the response time starts to increase exponentially in the unrefactored code. So, whith this parameters we can asume the importance of the refactoring code whith solid mesured results of efficiency.

In addition if we pay attention to the response time when we analyze the 2018 stocks, we can notice more clear difference between the refactored and the unrefactored code. It is not very clear for me the reasons of that but we can determine that the information we need to handle in the 2018 is more difficult or has more functions to be runned by our code because of the variability of the data. So to be more clear, this is why is very important the refactoring on the codes. If we handle complex data with so many variations, we will bring up issues during the processing of the code if we not make it clean or refactored.

## Advantages and disadvantages of refactoring code

Now that we talked about the anaylsis and results of the importance refactoring our codes, we can determina some pros and cons about if it is necessary to refactor our code or not. According to the information we collected with this project and the different performances we face working with both codes, we can establish the following conclusions:

**The advantages**

- Reduce the technical issues and easy to mantain. Refactoring our code will allow us to reduce the possibilities of have long time of responses related to a bad performance of the code. Also it will be esay to mantain in the way we can find issues or improve the steps to reach our goals.
- Easy read. With a cleaner code, it allow us to add or remove features. Other programmers can easily work on our code improving the way it works whithout spending time finding how our code is working. Also the beginners and the pros can work on our code if it is simple and well structured.
- Efficent response. It leads to improve easily the "architecture" of the code in a way to reduce the response time and to clean the proccess *per se*. We will help our program to work without complex actions to take making it simple and to have an analysis of large data in a very short time.
- Basis for many codes. If our refactored code is very clean and effcient it will be the basis of new codes. Simple as that.

**The disadvantages**

- Spending time. We need time to expend refactoring the code. We need to spend ours finding a way to make simple and cleaner our code. It sounds good the idea of clean our code but of course it needs time and habilities to improve our architectures. The most we work on the scripst the least time we spend.
- Expertise. Related to the first disadvantage, we need some kind of expertise to avoid the waste of time in a script that it is hard for us. If the code is not complex we can handle it, but what if our code is very complex one with a lot of functions to do.
- Get lost in the way. Spending to much time in refactoring our code, we can lose the idea of the goal. Even worst, we can change the code in a worst way missing things we want it and changing the course of our needs. Of course we can save the code before work on it, but there will be always the risk of lose information.

## Conclusion
In conclusion, an accoridng to the results gived to Steve(our customer), refactoring is a main step in the VBA coding script.The pros are more relevant than the cons: **KEEP IT SIMPLE!**. The value added with factoring our code is essencial on this kind of coding, we need to spend time cleaning the codes to improve the needs of our customers, also to invest the time on it will increase our vision on codes and make us clear thinkers. So it is a win to win. The disdavantages are thing we can work on, there are not breaking points to avoid the refactoring. All is about the effort we need to take on our code. So you need to refactore if you want to bring VBA script code solutions for stock analysis or any analysis of big data you are thinking.
