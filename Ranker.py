from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import time
import logging
from collections import defaultdict
import os
from os import listdir
from os.path import isfile, join
from io import StringIO
import statistics
from datetime import datetime

def Location():
    global location
    location = filedialog.askdirectory() + '/'
    if len(location) > 1:
        Debug.config(text="Click The Download Button if this is the first time.",fg="green")
        logging.basicConfig(filename=location+'log.txt', filemode='a', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')
    else:
        Debug.config(text="Pick a Folder!",fg="red")

def DownloadData():
    loadTime = int(input1.get())
    try:
        v = int(loadTime)
    except ValueError:
        Debug.config(text="Pick a Number for loadTime!",fg="red")
        return
    if int(loadTime) < 0:
        Debug.config(text="Pick a Whole Number for loadTime!",fg="red")
        return
    try:
        v = len(location)
    except Exception:
        Debug.config(text="Pick a Folder!",fg="red")
        return
    Debug.config(text="Scraping the website...",fg="white")

    try:
        driver = webdriver.Edge()
    except Exception as e:
        Debug.config(text="Unexpected exception. Downloading Drivers.",fg="red")
        logging.error("Driver Error", exc_info=True)
        try:
            service = Service(EdgeChromiumDriverManager().install())
            driver = webdriver.Edge(service=service)
            Debug.config(text="Scraping the website...",fg="white")
        except Exception as e:
            Debug.config(text="Fatal error. Unable to continue.",fg="red")
            logging.error("Fatal Error", exc_info=True)
            return
    driver.get("https://josaa.admissions.nic.in/Applicant/seatallotmentresult/currentorcr.aspx")
    driver.implicitly_wait(loadTime)

    links = []
    listA = "1", "2", "3", "4", "5"
    scroll = driver.execute_script("return window.pageYOffset;")

    for i in listA:
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[1]/div/div").click() # round number
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[1]/div/div/div/ul/li[" + str(int(i)+1) + "]").click()
        time.sleep(loadTime)
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[2]/div/div").click() # institute type
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[2]/div/div/div/ul/li[2]").click()
        time.sleep(loadTime)
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[3]/div/div").click() # institute name
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[3]/div/div/div/ul/li[2]").click()
        time.sleep(loadTime)
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[4]/div/div").click() # academic program
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[4]/div/div/div/ul/li[2]").click()
        time.sleep(loadTime)
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[5]/div/div").click() # category
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[5]/div/div/div/ul/li[2]").click()
        time.sleep(loadTime)
        driver.find_element(By.XPATH, "/html/body/form/div[3]/div[1]/div[6]").click() # submit
        time.sleep(loadTime)

        table = driver.find_element(By.XPATH, "/html/body/form/div[3]/div[2]/div[2]/div/div/table")
        driver.execute_script("arguments[0].scrollIntoView();", table)
        links.append(table.get_attribute("outerHTML"))
        driver.execute_script(f"window.scrollTo(0, {scroll});")
        time.sleep(loadTime)
        Debug.config(text="Scraped Round Number " + str(i) + "...",fg="white")

    driver.quit()
    i = 1
    for link in links:
        if os.path.exists(location + "Round " + str(i) + ".xlsx"):
            os.remove(location + "Round " + str(i) + ".xlsx")
        df = pd.read_html(StringIO(link))[0]
        df.to_excel(location + "Round " + str(i) + ".xlsx", index=False)
        Debug.config(text="Downloaded Round Number " + str(i) + "...",fg="white")
        i = i + 1
    Debug.config(text="Select preferred Institute and get their names.",fg="green")

def GetResults():
    try:
        v = len(location)
    except Exception:
        Debug.config(text="Pick a Folder!",fg="red")
        return
    if dropdown2.get() == "Select Round Number:":
        Debug.config(text="Pick a Round!",fg="red")
    rank = input2.get()
    try:
        rank = int(rank)
    except ValueError:
        Debug.config(text="Enter your rank!",fg="red")
        return
    if int(rank) < 0:
        Debug.config(text="A rank is a whole number...",fg="red")
        return
    Debug.config(text="Getting Results...",fg="white")

    t = dropdown1.get()
    college = dropdown3.get()
    course = dropdown4.get()
    quota = dropdown5.get()
    category = dropdown6.get()
    gender = dropdown7.get()
    
    df = pd.read_excel(location + "Round " + dropdown2.get() + ".xlsx")
    colleges = list(df.loc[1:]["Institute"])
    courses = list(df.loc[1:]["Academic Program Name"])
    quotas = list(df.loc[1:]["Quota"])
    categories = list(df.loc[1:]["Seat Type"])
    genders = list(df.loc[1:]["Gender"])
    crs = list(df.loc[1:]["Closing Rank"])

    my_dict = defaultdict(list)
    i = 0
    while i < len(crs):
        j = crs[i]
        if "." in j:
            j = j.split(".")[0]
        if "P" in j:
            j = j.split("P")[0]
        if "p" in j:
            j = j.split("p")[0]
        k = int(j)
        if rank <= k:
            if Check(i, college, course, quota, category, gender, colleges, courses, quotas, categories, genders) == True:
                if t == "All" and str("Indian Institute of Technology").lower() not in colleges[i].lower():
                    if var.get() == 0:
                        my_dict[colleges[i]].append(courses[i])
                    else:
                        my_dict[colleges[i]].append(str(courses[i]).split('(')[0].strip())
                elif t == "Government Funded Technical Institutions":
                    if str("Indian Institute of Information Technology").lower() not in colleges[i].lower() and str("National Institute of Technology").lower() not in colleges[i].lower() and str("Indian Institute of Technology").lower() not in colleges[i].lower():
                        if var.get() == 0:
                            my_dict[colleges[i]].append(courses[i])
                        else:
                            my_dict[colleges[i]].append(str(courses[i]).split('(')[0].strip())
                elif t == "Indian Institute of Information Technology" or t == "National Institute of Technology" or t == "Indian Institute of Technology":
                    if t.lower() in colleges[i].lower() and courses[i] not in my_dict[colleges[i]]:
                        if var.get() == 0:
                            my_dict[colleges[i]].append(courses[i])
                        else:
                            my_dict[colleges[i]].append(str(courses[i]).split('(')[0].strip())
        i = i + 1

    try:
        planck_length = max(len(v) for v in my_dict.values())
        data = {k: v + [None] * (planck_length - len(v)) for k,v in my_dict.items()}
        df = pd.DataFrame(data)
        df.to_excel(location + "Results " + str(GetIndex()) + ".xlsx", index=False)
        SaveData(location + "Results " + str(GetIndex()) + ".xlsx", "Success")
        Debug.config(text="Finished.",fg="green")
    except Exception as e:
        logging.error("Data Error", exc_info=True)
        SaveData(location + "Results " + str(GetIndex()) + ".xlsx", "Failure")

def Check(i, college, course, quota, category, gender, colleges, courses, quotas, categories, genders):
    t = False
    if college == "All":
        t = True
    elif colleges[i] == college:
        t = True
    else:
        return False
    if var.get() == 0:
        if course == "All":
            t = True
        elif courses[i] == course:
            t = True
        else:
            return False
    else:
        if course == "All":
            t = True
        elif str(courses[i]).split('(')[0].strip() == str(course).split('(')[0].strip():
            t = True
        else:
            return False
    if quotas[i] == "AI":
        t = True
    elif quotas[i] == quota:
        t = True
    elif varB.get() == 1:
        t = True
    else:
        return False
    if categories[i] == category:
        t = True
    else:
        return False
    if genders[i] == "Gender-Neutral":
        t = True
    elif genders[i] == gender:
        t = True
    else:
        return False
    return t

def GetValues():
    t = dropdown1.get()
    if t == "Select Institute type:":
        Debug.config(text="Pick the type of Institute!",fg="red")
        return
    try:
        v = len(location)
    except Exception:
        Debug.config(text="Pick a Folder!",fg="red")
        return
    Debug.config(text="Getting Values...",fg="white")

    df = pd.read_excel(location + "Round " + dropdown2.get() + ".xlsx")
    colleges = list(df.loc[1:]["Institute"].unique())
    courses = list(df.loc[1:]["Academic Program Name"].unique())
    quotas = list(df.loc[1:]["Quota"].unique())
    categories = list(df.loc[1:]["Seat Type"].unique())
    genders = list(df.loc[1:]["Gender"].unique())

    if dropdown1.get() != "All":
        if dropdown1.get() == "Government Funded Technical Institutions":
            for i in reversed(colleges):
                if str("Indian Institute of Information Technology").lower() in i.lower() or str("National Institute of Technology").lower() in i.lower() or str("Indian Institute of Technology").lower() in i.lower():
                    colleges.remove(i)
        else:
            for i in reversed(colleges):
                if dropdown1.get().lower() not in i.lower():
                    colleges.remove(i)
    
    if var.get() == 1:
        course = []
        for i in courses:
            course.append(str(i).split('(')[0].strip())
        course = [*{*course}] # remove duplicates
        course.insert(0, "All")

    colleges.insert(0, "All")
    courses.insert(0, "All")

    dropdown3['values'] = colleges
    if var.get() == 1:
        dropdown4['values'] = course
    else:
        dropdown4['values'] = courses
    dropdown5['values'] = quotas
    dropdown6['values'] = categories
    dropdown7['values'] = genders
    Debug.config(text="Select the values.",fg="green")

def Find(condition):
    try:
        v = len(location)
    except Exception:
        Debug.config(text="Pick a Folder!",fg="red")
        return
    Debug.config(text="Getting Data...",fg="white")
    ranks = GetData()
    if condition == "min":
        try:
            Debug.config(text="The minimum rank for given options is: " + str(max(ranks)),fg="green")
        except Exception as e:
            Debug.config(text="No data available for the given options.",fg="red")
            logging.error("Min Rank Error", exc_info=True)
    if condition == "av":
        try:
            Debug.config(text="The average rank for given options is: " + str(statistics.fmean((ranks))),fg="green")
        except Exception as e:
            Debug.config(text="No data available for the given options.",fg="red")
            logging.error("Av Rank Error", exc_info=True)

def GetData():
    ranks = []
    t = dropdown1.get()
    college = dropdown3.get()
    course = dropdown4.get()
    quota = dropdown5.get()
    category = dropdown6.get()
    gender = dropdown7.get()
    colleges = []
    courses = []
    quotas = []
    categories = [] 
    genders = []
    crs = []

    i = 1
    while i < 6:
        df = pd.read_excel(location + "Round " +  str(i) + ".xlsx")
        colleges.extend(list((df.loc[1:]["Institute"])))
        courses.extend(list((df.loc[1:]["Academic Program Name"])))
        quotas.extend(list((df.loc[1:]["Quota"])))
        categories.extend(list((df.loc[1:]["Seat Type"])))
        genders.extend(list((df.loc[1:]["Gender"])))
        crs.extend(list((df.loc[1:]["Closing Rank"])))
        i = i + 1

    i = 0
    while i < len(colleges):
        j = crs[i]
        if "." in j:
            j = j.split(".")[0]
        if "P" in j:
            j = j.split("P")[0]
        if "p" in j:
            j = j.split("p")[0]
        k = int(j)
        if dropdown3.get() == "Select Institute name:" or dropdown3.get() == "All":
            college = colleges[i]
        if dropdown4.get() == "Select Course:" or dropdown4.get() == "All":
            course = courses[i]
        if var.get() == 1:
            course = str(course).split('(')[0].strip()
        if dropdown5.get() == "Select Quota:":
            quota = quotas[i]
        if dropdown6.get() == "Select Category:":
            category = categories[i]
        if dropdown7.get() == "Select Gender:":
            gender = genders[i]
        if Check(i, college, course, quota, category, gender, colleges, courses, quotas, categories, genders) == True:
            if t == "All":
                ranks.append(k)
            elif t == "Government Funded Technical Institutions":
                if str("Indian Institute of Information Technology").lower() not in colleges[i].lower() and str("National Institute of Technology").lower() not in colleges[i].lower() and str("Indian Institute of Technology").lower() not in colleges[i].lower():
                    ranks.append(k)
            elif t == "Indian Institute of Information Technology" or t == "National Institute of Technology" or t == "Indian Institute of Technology":
                if t.lower() in colleges[i].lower():
                    ranks.append(k)
        i = i + 1
    return ranks

def GetIndex():
    i = 1
    onlyfiles = [f for f in listdir(location.replace('/', '')) if isfile(join(location.replace('/', ''), f))]
    for file in onlyfiles:
        if "Result" in file:
            i = i + 1
    return i

def SaveData(name, state):
    list = []
    list.append("Result " + str(GetIndex()) + ": " + str(datetime.now()) + " " + state)
    list.append(" Location - " + name)
    list.append(" Configurations:")
    list.append(" loadTime- " + str(input1.get()) + ", Institute type- " + dropdown1.get() + ", Round number- " + dropdown2.get() + ", Duplicate courses- " + str(var.get()))
    list.append(" Rank- " + str(input2.get()) + ", Institute name- " + dropdown3.get() + ", Course- " + dropdown4.get() + ", Quota- " + dropdown5.get() + ", Category- " + dropdown6.get())
    list.append(" Gender- " + dropdown7.get() + ", Ignore Quota- " + str(varB.get()) + "\n")
    with open(location + "save.txt", 'a') as f:
        f.writelines(list)
        f.close()

def on_1_click(event):
    if input1.get() == "How long should the app wait for loading?":
        input1.delete(0, tk.END)
        input1.config(fg='black')

def on_2_click(event):
    if input2.get() == "What is your Rank?":
        input2.delete(0, tk.END)
        input2.config(fg='black')

app = tk.Tk()
app.title("Ranker")
app.configure(bg='black')

title = tk.Label(app, text="Ranker", font=("Helvetica", 25), fg = "gold", bg = "black")
title.pack(pady=10)

input1 = tk.Entry(app, width=150, fg='black', font=('Helvetica', 10))
input1.insert(0, "How long should the app wait for loading?")
input1.bind('<FocusIn>', on_1_click)
input1.pack(pady=5)

button1 = tk.Button(app, text="Choose Location", fg = "blue", command=Location)
button1.pack(pady=10)

button2 = tk.Button(app, text="Download data", fg = "blue", command=DownloadData)
button2.pack(pady=10)

dropdown1 = ttk.Combobox(app, width = 150, values=["All", "Government Funded Technical Institutions", "Indian Institute of Information Technology", "National Institute of Technology", "Indian Institute of Technology"]) # types
dropdown1.set("Select Institute type:")
dropdown1.bind("<Key>", lambda e: "break")
dropdown1.pack(pady=5)

dropdown2 = ttk.Combobox(app, width = 150, values=["1", "2", "3", "4", "5"]) # rounds
dropdown2.set("Select Round Number:")
dropdown2.bind("<Key>", lambda e: "break")
dropdown2.pack(pady=5)

var = tk.IntVar()
c1 = tk.Checkbutton(app, text = "Try and remove duplicate courses?", variable=var, onvalue=1, offvalue=0)
c1.pack(pady=5)

button3 = tk.Button(app, text="Get Values", fg = "blue", command=GetValues)
button3.pack(pady=10)

input2 = tk.Entry(app, width=150, fg='black', font=('Helvetica', 10))
input2.insert(0, "What is your Rank?")
input2.bind('<FocusIn>', on_2_click)
input2.pack(pady=5)

dropdown3 = ttk.Combobox(app, width = 150, values=[]) # names
dropdown3.set("Select Institute name:")
dropdown3.bind("<Key>", lambda e: "break")
dropdown3.pack(pady=5)

dropdown4 = ttk.Combobox(app, width = 150, values=[]) # courses
dropdown4.set("Select Course:")
dropdown4.bind("<Key>", lambda e: "break")
dropdown4.pack(pady=5)

dropdown5 = ttk.Combobox(app, width = 150, values=[]) # quota
dropdown5.set("Select Quota:")
dropdown5.bind("<Key>", lambda e: "break")
dropdown5.pack(pady=5)

dropdown6 = ttk.Combobox(app, width = 150, values=[]) # category
dropdown6.set("Select Category:")
dropdown6.bind("<Key>", lambda e: "break")
dropdown6.pack(pady=5)

dropdown7 = ttk.Combobox(app, width = 150, values=[]) # gender
dropdown7.set("Select Gender:")
dropdown7.bind("<Key>", lambda e: "break")
dropdown7.pack(pady=5)

varB = tk.IntVar()
c2 = tk.Checkbutton(app, text = "Ignore Quota?", variable=varB, onvalue=1, offvalue=0)
c2.pack(pady=5)

button4 = tk.Button(app, text="Get Results", fg = "blue", command=GetResults)
button4.pack(pady=10)

frame1 = tk.Frame(app)
frame1.pack(pady=10)

button5 = tk.Button(frame1, text="Find minimum rank for given options", fg = "blue", command=lambda: Find("min"))
button5.grid(row=0, column=0, padx=10)

button6 = tk.Button(frame1, text="Find average rank for given options", fg = "blue", command=lambda: Find("av"))
button6.grid(row=0, column=1, padx=10)

Debug = tk.Label(app, text="Start with loading time and download location.", font=("Helvetica", 10), bg = "black", fg = "white")
Debug.pack(pady=10)

app.mainloop()
