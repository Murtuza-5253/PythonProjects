


# ----------------------------------------  IMPORTING LIBRARIES ----------------------------------------------------------------
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import xlwings as xw 

user_input = int(input('''
                        Enter 1 to see the analysis of Reviews given by Customers
                        Enter 2 to see the analysis of different payment methods used by the Customers
                        Enter 3 to see the analysis of Top Consumer States of India
                        Enter 4 to see the analysis of Top Consumer Cities of India
                        Enter 5 to see the analysis of Top Selling Product Categories
                        Enter 6 to see the analysis of Reviews for All Product Categories
                        Enter 7 to see the analysis of Number of Orders Per Month Per Year
                        Enter 8 to see the analysis of Reviews for Number of Orders Per Month Per Year
                        Enter 9 to see the analysis of Number of Orders Across Parts of a Day
                        Enter 10 to see the Full Report in notebook
                        Enter 11 to see the Full Report in excel
                        '''))

# ----------------------------------------   READING THE DATASETS   -------------------------------------------------------------

#reading the review dataset
df_review = pd.read_csv('C:\\Users\\Murtuza pipulyawala\\Desktop\\review_dataset.csv')

#reading order dataset
df_order = pd.read_csv('C:\\Users\\Murtuza pipulyawala\\Desktop\\orders_2020_2021_DataSet_Updated.csv',parse_dates=[1])


# --------------------------------- 1) ANALYSIS OF REVIEWS GIVEN BY CUSTOMER ----------------------------------------------------
def review_by_cust():
    #dropping NaN values and saving in another variable
    df_review_without_null = df_review.dropna()

#plotting the graph
    fig = plt.figure()
    plt.title('ANALYSIS OF REVIEWS GIVEN BY CUSTOMER',fontsize=18)
    review_count = df_review_without_null['stars'].value_counts()
    plt.subplot(3,4,2)
    review_count.plot(kind='bar',label='<2.0 is negative rating',figsize=(40,20))
    plt.ylabel('Number of reviews',fontsize=22)
    plt.xlabel('Reviews')
    plt.show()

    

# ---------------------------------- 2) Analysis of different payment methods used by the customer -----------------------------   
def diff_pay_met():
    #seperating the Payment method column
    #dropping NaN values
    #splitting a string into list and grabbing the first element
    #converting to dataframe
    df_order_payment_method = df_order['Payment Method']
    df_order_payment_method = df_order_payment_method.dropna()
    df_order_payment_method = df_order_payment_method.str.split().apply(lambda a : a[0])
    df_order_payment_method = df_order_payment_method.to_frame()

    #plotting the graph
    fig = plt.figure()
    plt.title('Analysis of different payment methods used by the customer',fontsize=18)
    payment_met = df_order_payment_method['Payment Method'].value_counts()
    payment_met.plot(kind='pie', autopct='%0.2f%%',figsize=(8,4))
    plt.show()



# ------------------------------------- 3) Analysis of top consumer states of India --------------------------------------------   
def top_states():
    #seperating the Shipping State column
    #dropping NaN values
    #converting to dataframe
    #keeping only string which starts with IN
    df_order_shipping_state = df_order['Shipping State']
    df_order_shipping_state = df_order_shipping_state.to_frame()
    df_order_shipping_state = df_order_shipping_state.dropna()
    keep=['IN-']
    df_order_shipping_state = df_order_shipping_state[df_order_shipping_state['Shipping State'].str.contains('|'.join(keep))]

    #plotting the graph
    fig = plt.figure()
    plt.title('Analysis of top consumer states of India',fontsize=18)
    top_states = df_order_shipping_state['Shipping State'].value_counts().iloc[:10]
    top_states.plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    plt.show()

# ---------------------------------------- 4) Analysis of top consumer cities of India -----------------------------------------
def top_cities():
    #seperating the Shipping cities column
    #converting to dataframe
    df_order_shipping_cities = df_order['Shipping City']
    df_order_shipping_cities = df_order_shipping_cities.to_frame()

    #plotting the graph
    fig = plt.figure()
    plt.title('Analysis of top consumer cities of India',fontsize=18)
    df_order_shipping_cities = df_order_shipping_cities['Shipping City'].value_counts().iloc[:7]
    df_order_shipping_cities.plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')

# ------------------------------------------5) Analysis of top selling product categories --------------------------------------
def top_sell_prod():
    #seperating the category column
    #converting to dataframe
    df_sell_product_category = df_review['category']
    df_sell_product_category = df_sell_product_category.to_frame()
    
#plotting the graph
    fig = plt.figure()
    plt.title('Analysis of top selling products',fontsize=18)
    sell_prod_cat = df_sell_product_category['category'].value_counts().iloc[:10]
    sell_prod_cat.plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    plt.show()


# ----------------------------------------- 6) Analysis of reviews for all product categories ----------------------------------
def review_sell_prod():

    
#plotting the graph
    fig = plt.figure()
    plt.title('Analysis of reviews for all product categories',fontsize=18)
    df_review.groupby('category')['stars'].count().plot(kind='bar', figsize=(40, 50))
    plt.show()

    
 # ------------------------------------------ 7) Analysis of number of order per month per year ------------------------------
def order_PmPy():
    df_order_date_PmPy = df_order['Fulfillment Date and Time Stamp']
    df_order_date_PmPy =df_order_date_PmPy.to_frame()
    df_order_date_PmPy = df_order_date_PmPy.dropna()
    df_order_date_PmPy['Year'] = df_order_date_PmPy['Fulfillment Date and Time Stamp'].dt.year
    df_order_date_PmPy['Month'] = df_order_date_PmPy['Fulfillment Date and Time Stamp'].dt.month
    df_order_date_PmPy['Day'] = df_order_date_PmPy['Fulfillment Date and Time Stamp'].dt.day
    df_order_date_PmPy = df_order_date_PmPy.drop(['Fulfillment Date and Time Stamp'],axis=1)
    df_order_date_PmPy = df_order_date_PmPy.groupby(['Year','Month']).count()
    
    df_order_date_PmPy.loc[2016].plot(kind='bar', figsize=(10, 8))
    plt.title('2016')


    df_order_date_PmPy.loc[2017].plot(kind='bar', figsize=(10, 8))
    plt.title('2017')

    
    df_order_date_PmPy.loc[2018].plot(kind='bar', figsize=(10, 8))
    plt.title('2018')

        
    df_order_date_PmPy.loc[2019].plot(kind='bar', figsize=(10, 8))
    plt.title('2019')

    
    df_order_date_PmPy.loc[2020].plot(kind='bar', figsize=(10, 8))
    plt.title('2020')

    
    df_order_date_PmPy.loc[2021].plot(kind='bar', figsize=(10, 8))
    plt.title('2021')
    plt.show()
    
# ---------------------------- 8)Analysis of reviews for number of orders per month per year -----------------------------------
def rev_PmPy():    
    df_rev_PmPy = pd.concat([df_order,df_review['stars']],axis=1)
    df_rev_PmPy = df_rev_PmPy[['Fulfillment Date and Time Stamp','stars']]
    df_rev_PmPy = df_rev_PmPy.dropna()
    
#plotting the graph
    fig = plt.figure()
    plt.title('reviews for number of orders per month per year',fontsize=18)
    df_review['month'] = df_order['Fulfillment Date and Time Stamp'].dt.month
    rev_pmpy = df_review.groupby('month')['stars'].value_counts()
    rev_pmpy.plot(kind='bar',label='')
    plt.ylabel('number of ratings',fontsize=15)
    plt.xlabel('review ratings per month per year')
    plt.show()

    
# ------------------------------- 9) Analysis of number of orders across parts of day ------------------------------------------
def parts_of_day():
    df_order_parts_of_day = df_order['Fulfillment Date and Time Stamp']
    df_order_parts_of_day = df_order_parts_of_day.to_frame()
    df_order_parts_of_day = df_order_parts_of_day.dropna()
    df_order_parts_of_day['Day'] = df_order_parts_of_day['Fulfillment Date and Time Stamp'].dt.day
    df_order_parts_of_day['Time_in_hours'] = df_order_parts_of_day['Fulfillment Date and Time Stamp'].dt.strftime('%H')
    df_order_parts_of_day = df_order_parts_of_day.drop(['Fulfillment Date and Time Stamp'],axis=1)
    df_order_parts_of_day = df_order_parts_of_day['Time_in_hours'].astype(int)
    df_order_parts_of_day = df_order_parts_of_day.to_frame()

    plt.subplot(1,4,1)
    df_order_parts_of_day[df_order_parts_of_day['Time_in_hours'].between(5,12,inclusive='both')].value_counts().plot(kind='pie', figsize=(20, 16),autopct='%0.2f%%')
    plt.title('Morning Timings')
    
    plt.subplot(1,4,2)
    df_order_parts_of_day[df_order_parts_of_day['Time_in_hours'].between(13,16,inclusive='both')].value_counts().plot(kind='pie', figsize=(20, 16),autopct='%0.2f%%')
    plt.title('Afternoon Timings')
    
    plt.subplot(1,4,3)     
    df_order_parts_of_day[df_order_parts_of_day['Time_in_hours'].between(16,23,inclusive='both')].value_counts().plot(kind='pie', figsize=(20, 16),autopct='%0.2f%%')
    plt.title('Evening Timings')
    
    plt.subplot(1,4,4)
    df_order_parts_of_day[df_order_parts_of_day['Time_in_hours'].between(0,3,inclusive='both')].value_counts().plot(kind='pie', figsize=(20, 16),autopct='%0.2f%%')
    plt.title('Night Timings')

    plt.show()
    
#--------------------------------------------------- full report ----------------------------------------------------------------
def Full_Rep():
    df_order = pd.read_csv('C:\\Users\\Murtuza pipulyawala\\Desktop\\orders_2020_2021_DataSet_Updated.csv')
    df_review = pd.read_csv('C:\\Users\\Murtuza pipulyawala\\Desktop\\review_dataset.csv')
    fig = plt.figure()
    plt.subplot(3, 4, 1)
    plt.title('Review Analysis Given by Constumer', fontsize=20)
    review_count = df_review['stars'].value_counts()
    review_count.plot(kind='bar', label='<2.0 is negative rating')
    plt.ylabel("Number of Reviews", fontsize=20)
    plt.xlabel("Reviews", fontsize=5)
    
    plt.subplot(3, 4, 2)
    plt.title('Payment method Used By Customers', fontsize=20)
    df_order['Payment Method'].dropna().str.split().apply(lambda x: x[0]).value_counts().plot(kind='pie', autopct='%0.2f%%',
                                                                                         figsize=(8, 4))
    plt.subplot(3, 4, 3)
    plt.title('Top Consumer States of India', fontsize=20)
    df_order["Billing State"].dropna().value_counts().head().plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    
    plt.subplot(3, 4, 4)
    plt.title('Top Consumer Cities of India', fontsize=20)
    df_order["Billing City"].dropna().value_counts().head().plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    
    plt.subplot(3, 4, 5)
    plt.title('Top Selling Product Categories', fontsize=20)
    df_review["category"].value_counts().head(10).plot(kind='pie', figsize=(10, 5), autopct='%0.2f%%')
    
    plt.subplot(3, 4, 6)
    plt.title('Reviews for All Product Categories', fontsize=20)
    df_review.groupby('category')['stars'].count().plot(kind='bar', figsize=(40, 50))
    
    plt.subplot(3, 4, 7)
    plt.title('Number of Orders Per Month Per Year', fontsize=20)
    pd.to_datetime(df_order['Fulfillment Date and Time Stamp']).dt.month.value_counts().plot(kind='pie', autopct='%0.2f%%',
                                                                                  shadow=True, figsize=(8, 4))
    plt.subplot(3, 4, 8)
    plt.title('Reviews for Number of Orders Per Month Per Year', fontsize=20)
    df_review['month'] = pd.to_datetime(df_order['Fulfillment Date and Time Stamp']).dt.month
    df_review.groupby("month")['stars'].value_counts().plot(kind='bar', figsize=(40, 50))
    
    plt.subplot(3, 4, 9)
    plt.title('Number of Orders Across Parts of a Day', fontsize=20)
    orders = pd.to_datetime(df_order['Fulfillment Date and Time Stamp']).dt.strftime('%H:%M:%S').value_counts().values
    plt.plot(orders)
    
    
    wb = xw.Book()
    sht=wb.sheets[0]
    sht.name = "excel charts"
    sht.pictures.add(fig, name="Excel Charts", update=True, left=sht.range("A4").left, top=sht.range("A4").top,
                     height=1000, width=1500)
    
#conditions
if user_input == 1:
     review_by_cust()
        
if user_input == 2:
    diff_pay_met()
    
if user_input == 3:
    top_states()
    
if user_input == 4:
    top_cities()
    
if user_input == 5:
    top_sell_prod()

if user_input == 6:
    review_sell_prod()
    
if user_input == 7:
    order_PmPy()
    
if user_input == 8:
    rev_PmPy()

if user_input == 9:
    parts_of_day()
    
if user_input == 10:
    review_by_cust()
    diff_pay_met()
    top_states()
    top_cities()
    top_sell_prod()
    review_sell_prod()
    order_PmPy()
    rev_PmPy()
    parts_of_day()
    
if user_input == 11:
    Full_Rep()





