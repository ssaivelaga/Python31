#import libraries
import pandas as pd
import matplotlib as mlt
import matplotlib.pyplot as plt
from tkinter import *
import tkinter.ttk as ttk
import functools
import prediction as pdt

#global variable
change_index = 0

#This function analyze the companies data
def self_operations(filename="",sheet=0,ncol=1,index="",title="",xlab="Year",ylab="Price (In Million $)"):

    #create dataframe from excel file 
    df = pd.read_excel(filename,sheet_name=sheet,header=0,index_col = False)

    #invert dataframe
    df = df.set_index(index).T.rename_axis(xlab).rename_axis(None,1)
    
    #plot dataframe
    ax = df.plot.bar(rot=0,figsize=(8,6))

        
    #set x axis ticks
    #plt.xticks(range(2010, 2018))

    #add legend on graph and set its position
    legend = plt.legend(loc="upper left",fontsize = 'xx-small', bbox_to_anchor=(0, 1),ncol=ncol)
    legend.draggable()

    #set scale to auto to adjust according to axis values
    plt.autoscale(tight=False)

    #set y axis label
    plt.ylabel(ylab)

    #set figure title
    plt.title(title)

    #display the plotted graph
    plt.show()
#END

#This method plots a graph to compare two companies
def compare(sheet=0,index="",ylabel="",title=""):

    #create list of company names
    list = ['NCR','CoBiz','Fiserv','Cathay']

    #create dataframe from excel file
    df = pd.read_excel('Comparison.xlsx',sheet_name=sheet,header=0,index_col = index)

    #if the selected item in cmb_company1 is before the selected item in cmb_company2
    if change_index <= cmb_company2.current():

        #filter database to select only two columns
        df = df[[list[cmb_company1.current()],list[1+cmb_company2.current()]]]

    #if the selected item in cmb_company1 comes after the selected item in cmb_company2
    else:
        #filter database to select only two columns        
        df = df[[list[cmb_company1.current()],list[cmb_company2.current()]]]

    #plot bar graph on dataframe
    df.plot.bar(rot=0,figsize=(8,6))

    #add legend on graph and set its position
    legend = plt.legend(loc="upper left",bbox_to_anchor=(0, 1),ncol=1)
    legend.draggable()

    #set scale to auto to adjust according to axis values
    plt.autoscale(tight=False)

    #set y axis label
    plt.ylabel(ylabel)

    #set title of figure
    plt.title(title)

    #display the plotted graph        
    plt.show()
#END


#This method plots a graph to compare all companies
def compare_all(sheet=0,index="",ylabel="",title=""):

    #create dataframe from excel file
    df = pd.read_excel('Comparison.xlsx',sheet_name=sheet,header=0,index_col = index)

    #plot bar graph on dataframe
    df.plot.bar(rot=0,figsize=(8,6))

    #add legend on graph and set its position
    legend = plt.legend(loc="upper left",bbox_to_anchor=(0, 1),ncol=1)
    legend.draggable()

    #set scale to auto to adjust according to axis values
    plt.autoscale(tight=False)

    #set y axis label
    plt.ylabel(ylabel)

    #set title of figure
    plt.title(title)

    #display the plotted graph        
    plt.show()
#END


#This function plots the graph on table data
def plot_graph():

    #create list of file paths
    cmp_filenames = ['NCR Financial Data.xlsx','CoBiZ Financial Data.xlsx','Fiserv, Inc Financial Data.xlsx','CATHAY Financial Data.xlsx']

    #create list of index columns for different companies
    index = ['In millions','In Thousands','In millions','In Thousands']

    #if cmb_company's selected value is "CoBiz Financial Inc."
    if cmb_company.get() == "CoBiz Financial Inc.":
        #if cmb_function's selected value is "Consolidated Balance Sheet"
        if cmb_function.get() == "Consolidated Balance Sheet":
            #reset index list
            index = ['In millions','In millions','In millions','In Thousands']
        
    #plot and display a bar graph   
    self_operations(filename=cmp_filenames[cmb_company.current()],sheet=cmb_function.current(),ncol=1,index=index[cmb_company.current()],title=cmb_function.get())
#END


#This function display a table in different window and provide a button for plot
def display_table(df,title=""):

    #create new window
    window=Tk()
    #set window size
    window.resizable(0,0)
    #set window title
    window.title(title)

    #for each column in the header
    for i in range(len(df.columns)):

        #create a Entry field and set its size adjustable to text
        txt = Entry(window,font=('times', 13, 'bold'),relief=RIDGE,width=2+round(0.85*len(str(df.columns[i]))))

        #add Entry field to grid layout to create a table
        txt.grid(row=0, column=i, sticky=NSEW,padx=2,pady=2)

        #insert text from dataframe
        txt.insert(END, '%s' % df.columns[i])

        #set Entry state to read only
        txt.config(state="readonly")
        
    #for each row
    for i in range(df.shape[0]):

        #for each column in row
        for j in range(df.shape[1]):

            #create a Entry field and set its size adjustable to text
            txt = Entry(window,font=('times', 13, 'normal'),relief=RIDGE,width=2+round(0.85*len(str(df.iloc[i, j]))))

            #add Entry field to grid layout to create a table
            txt.grid(row=i+1, column=j, sticky=NSEW,padx=2,pady=2)

            #insert text from dataframe
            txt.insert(END, '%s' % df.iloc[i, j])

            #set Entry state to read only
            txt.config(state="readonly")
            
            
    #create a button for plotting
    btn = Button(window,text='Plot Graph', command=plot_graph)

    #add button to grid layout
    btn.grid(columnspan=len(df.columns))

    #loop the GUI   
    mainloop()
#END
    
    

#This function creates the main GUI window
def GUI():

    #create a window and set its properties
    window = Tk()
    window.geometry("650x360+400+20")
    window.resizable(0,0)
    window.title("Company Data Analyzer")


    #create a label for main title and set its properties
    title = Label(window,text='Company Data Analyser')
    title.place(x=170,y=10,height=30,width=300)
    title.config(font=('times', 20, 'bold'))
    

    ##create a label for sub title and set its properties
    title_single = Label(window,text='Single Company')
    title_single.place(x=80,y=65,height=30,width=150)
    title_single.config(font=('times', 14, 'bold'))

    #create a label for combo box and set its properties
    lbl_company = Label(window,text='Company')
    lbl_company.place(x=20,y=135,height=30,width=100)
    lbl_company.config(font=('times', 13, 'normal'))

    #create a combo box for selecting company
    global cmb_company
    cmb_company = ttk.Combobox(window, values=("NCR Corporation","CoBiz Financial Inc.","Fiserv Inc Financial Data","CATHAY Financial Data"))
    cmb_company.place(x=170,y=140,height=20,width=120)
    cmb_company.current(0)

    #create a label for combo box and set its properties
    lbl_function = Label(window,text='Select an Option')
    lbl_function.place(x=20,y=195,height=30,width=150)
    lbl_function.config(font=('times', 13, 'normal'))

    #create a combo box for selecting functionality
    global cmb_function
    cmb_function = ttk.Combobox(window, values=("Financial Data","Financial Conditions","Comprehensive Income","Consolidated Balance Sheet","Intangible assets"))
    cmb_function.place(x=170,y=200,height=20,width=120)
    cmb_function.current(0)


    #This function handles the button click event
    def btnclick():

        #create a list of file paths
        cmp_filenames = ['NCR Financial Data.xlsx','CoBiZ Financial Data.xlsx','Fiserv, Inc Financial Data.xlsx','CATHAY Financial Data.xlsx']
        #cmp_title = ['NCR-','CoBiz-','Fiserv-','CATHAY-']
            
        #create dataframe from excel file 
        df = pd.read_excel(cmp_filenames[cmb_company.current()],sheet_name=cmb_function.current(),header=0,index_col = False)

        #title = cmp_title[cmb_company.current()] + cmb_function.get()

        #display table in new window
        display_table(df,title=cmb_function.get())
    #END


    #create a button and add command to event handler function
    btn_single = Button(window,text='Submit',command=btnclick)
    btn_single.place(x=110,y=265,height=25,width=80)


    #This function handles the combo box Item change event
    def cmb_select(event):

        #create a list of company names to display in combo box
        cmp_list = ["NCR Corporation","CoBiz Financial Inc.","Fiserv Inc Financial Data","CATHAY Financial Data"]
        
        #for i in range 0 to 3
        for i in range(4):

            # if company name in list matches with selected value in combo box
            if cmp_list[i] == cmb_company1.get():

                #create a global variable
                global change_index

                #assign index value to variable
                change_index=i

                #remove matched item
                cmp_list.pop(i)

                #break the loop
                break
        #set values in combo box
        cmb_company2['values']=cmp_list

        #select current selected index
        cmb_company2.current(0)
    #END


    #create a Label and set subtitle of window
    title_multi = Label(window,text='Multiple Companies')
    title_multi.place(x=390,y=65,height=30,width=160)
    title_multi.config(font=('times', 14, 'bold'))

    #create a label for combo box and set its Properties
    lbl_company1 = Label(window,text='Company1')
    lbl_company1.place(x=340,y=110,height=30,width=100)
    lbl_company1.config(font=('times', 13, 'normal'))

    #create combo box for selecting company and  set its properties
    global cmb_company1
    cmb_company1 = ttk.Combobox(window, values=("NCR Corporation","CoBiz Financial Inc.","Fiserv Inc Financial Data","CATHAY Financial Data"))
    cmb_company1.place(x=490,y=115,height=20,width=120)
    cmb_company1.current(0)

    #bind the cmb_select function to combo box to call function on item change
    cmb_company1.bind("<<ComboboxSelected>>", cmb_select)

    #create a label for combo box and set its Properties
    lbl_company2 = Label(window,text='Company2')
    lbl_company2.place(x=340,y=160,height=30,width=100)
    lbl_company2.config(font=('times', 13, 'normal'))

    #create another combo box for selecting company and  set its properties
    global cmb_company2
    cmb_company2 = ttk.Combobox(window, values=("CoBiz Financial Inc.","Fiserv Inc Financial Data","CATHAY Financial Data"))
    cmb_company2.place(x=490,y=165,height=20,width=120)
    cmb_company2.current(0)

    #create a label for combo box and set its Properties
    lbl_function2 = Label(window,text='Select an Option')
    lbl_function2.place(x=340,y=210,height=30,width=150)
    lbl_function2.config(font=('times', 13, 'normal'))

    #create combo box for selecting functionality and  set its properties
    cmb_function2 = ttk.Combobox(window, values=("Financial Ratios","Returns"))
    cmb_function2.place(x=490,y=215,height=20,width=120)
    cmb_function2.current(0)


    #This function handles the button click event
    def btnclick2():

        #if selected value is equal to "Financial Ratios"
        if cmb_function2.get() == "Financial Ratios":

            #plot graph
            compare(sheet=0,index="Ratio",ylabel="Ratio",title="Comparison of Financial Ratios")

        #if selected value is equal to "Returns"
        elif cmb_function2.get() == "Returns":

            #plot graph
            compare(sheet=1,index="Years",ylabel="Price (In $)",title="Comparison of Cumulative Five Year Total Return")
    #END

    
    #create a button and set the command
    btn_multi = Button(window,text='Submit',command=btnclick2)

    #set size and location of button 
    btn_multi.place(x=440,y=265,height=25,width=80)


    #This function handles the button click event and creates a gui window
    def  btnclick3():

        #create a window and set its properties
        window = Tk()
        window.geometry("400x180+400+20")
        window.resizable(0,0)
        window.title("Compare all Companies")

        #create a label for title and set its properties
        title = Label(window,text='Company Data Analyser')
        title.place(x=50,y=10,height=30,width=300)
        title.config(font=('times', 20, 'bold'))

        
        #create a label for combo box and set its properties
        lbl_function = Label(window,text='Select an Option')
        lbl_function.place(x=65,y=80,height=30,width=120)
        lbl_function.config(font=('times', 13, 'normal'))

        #create a global combo box for selecting functionality
        global cmb_new
        cmb_new = ttk.Combobox(window, values=("Financial Ratios","Returns"))
        cmb_new.place(x=200,y=85,height=25,width=120)
        cmb_new.current(0)


        #This function handles the button click event
        def btn_click():

            #if selected value is equal to "Financial Ratios"
            if cmb_new.get() == "Financial Ratios":

                #plot graph
                compare_all(sheet=0,index="Ratio",ylabel="Ratio",title="Comparison of Financial Ratios")

            #if selected value is equal to "Returns"
            elif cmb_new.get() == "Returns":

                #plot graph
                compare_all(sheet=1,index="Years",ylabel="Price (In $)",title="Comparison of Cumulative Five Year Total Return")
        #END

        #create a button and set the command
        btn_single = Button(window,text='Submit',command=btn_click)

        #set size and location of button 
        btn_single.place(x=140,y=135,height=30,width=120)

    #END

    #This function handles the button click event and creates a gui window
    def  btnclick4():

        #plot predictive model
        pdt.subplot()


    #create a button and set the command
    btn_compareall = Button(window,text='Predictive model',command=btnclick4)

    #set size and location of button
    btn_compareall.place(x=160,y=320,height=30,width=150)

    
    #create a button and set the command
    btn_compareall = Button(window,text='Compare all Companies',command=btnclick3)

    #set size and location of button
    btn_compareall.place(x=340,y=320,height=30,width=150)

    #loop the GUI
    window.mainloop()
#END
    

GUI()


