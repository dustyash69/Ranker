# What is Ranker? Where Can I download it?
Ranker is a free college predictor for JEE Mains & Advanced, based on rank.
You can download the .exe version of this code [here](link).
The above link is of Google Drive. Click on "Download Anyway" when Google Drive asks to. The source code is available in `Ranker.py`.
After downloading the .exe file from the link above, open it. Alternatively, you can run the code from `Ranker.py` yourself. Apart from the used modules, you will need `beautifulsoup4, lxml & openpyxl`. Or, you can use `pip install pandas[all]`.

**Important Note - Microsoft Edge is required to use Ranker.**

# How do I use Ranker?

## If you're using Ranker for the first time:
1) In the `How long should the app wait for loading?` input field, enter a whole number. This number is the amount of seconds that the app waits for the page to load. This number depends on how fast your PC and Internet is. **3** usually does the trick.
2) Now, go to your File Manager and create a new folder where Ranker will store files. Click on the `Choose Location` button and select the Folder you just created.
3) Click on the `Download data` button. This will open a new tab in Microsoft Edge. Please do **NOT** interact with this window in any way. Let it do its job, undisturbed. This may take a while.

## If you've already downloaded the files before:
Then, Click on the `Choose Location` button and select the Folder you used before.

## What do I do now?
### For JEE Mains:
In the `Select Institute type:` dropdown, select any option except for `IIT`. `All` will select all types of Institutes, except for `IIT`.
### For JEE Advanced:
In the `Select Institute type:` dropdown, select `Indian Institute of Technology`.

1) Now, select any round from `Select Round Number:`. Colleges and Branches vary at each Rank in each Round.
2) There is a checkbox for `Try and remove duplicate courses?`. If you tick it, the app will try to find duplicate branches and remove them.
3) Then, click on the `Get Values` button. Wait for the values to load.
 
### For JEE Mains:
4) In the `What is your Rank?` Input field, Enter the rank you got in `JEE Mains`. **If you belong to Open Category, enter your CRL Rank otherwise enter your category rank.**
### For JEE Advanced:
4) In the `What is your Rank?` Input field, Enter the rank you got in `JEE Advanced`. **If you belong to Open Category, enter your CRL Rank otherwise enter your category rank.**
 
5) In the `Select Institute name:` dropdown, select the Institute you prefer. Select `All` if you do not have any preference.
6) In the `Select Course name:` dropdown, select the Branch you prefer. Select `All` if you do not have any preference.
7) In the `Select Quota:` dropdown, select your Quota. If you have a home state quota, please select that. This will also include `AI`, i.e., "All India" quota colleges. If you do not have any, select `AI`.
8) In the `Select Category:` dropdown, select your Category.
9) In the `Select Gender:` dropdown, select your gender. For men, select `Gender-neutral` otherwise select `Female only(including Supernumerary)`.
10) Click on the `Get Results` button. Wait for the app to compile your results.

## Where are the results?
The college and the branches you got at them are stored in an excel file in the folder you selected before. There is also a text file in it to guide you.

# Some Cool Stuff you can do in Ranker:
Ranker lets you find the **minimum and average rank required to get into any institute at any branch** for any category, gender, etc.
You can provide as little or as much information you want. Ranker will work with anything.
You can find the result of your query in the green text at the bottom of the app.
