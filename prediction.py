#import libraries
from sklearn.naive_bayes import GaussianNB
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
    

#This function plot a line chart
def predictive_model( filename,sheet,ax,title,income="Income (In million $)"):

        #load data in dataframe
        df= pd.read_excel(filename,sheet_name=sheet)

        #train data
        train_data = df.sample(frac=0.97,random_state=1)

        #test data
        test_data = df.loc[~df.index.isin(train_data.index)]

        #test set
        test_set = test_data[income].tolist()

        #column list
        columns = df.columns.tolist()

        #get object of naive byes
        gnb = GaussianNB()

        #train data
        gnb.fit(train_data[columns], train_data[income])

        #predict data
        predicted_data = gnb.predict(test_data[columns])

        #Plot graph
        #set title
        plt.title(title)
        #set y label
        plt.ylabel(income)
        #set x label
        plt.xlabel('year (2011 - 2017)')

        #set labels for x axis
        plt.xticks(np.arange(7), ('2011', '2012', '2013', '2014', '2015','2016','2017'))

        #plot test data
        ax.plot(test_set[:6], label='Train Data' ,color='yellow')
        #plot predicted data
        ax.plot(predicted_data[:6], label='Prediction',color='green')  

        #add legend
        plt.legend()
#END

#This function display all charts in one figure
def subplot():

    #set figure size
    plt.figure(figsize=(10,6))

    #NCR
    ax1 = plt.subplot2grid((2, 2), (0, 0))
    predictive_model('prediction.xlsx',0, ax1,"NCR")

    #CoBiz
    ax2 = plt.subplot2grid((2, 2), (0, 1))
    predictive_model('prediction.xlsx',1, ax2,"CoBiz",income="Income (In thousand $)")

    #Fiserv
    ax3 = plt.subplot2grid((2, 2), (1, 0))
    predictive_model('prediction.xlsx',2,ax3,"Fiserv")

    #Cathay
    ax4 = plt.subplot2grid((2, 2), (1, 1))
    predictive_model('prediction.xlsx',3, ax4,"Cathay","Income (In thousand $)")

    #compress charts
    plt.tight_layout()

    #display all charts in subplot
    plt.show()

#END


