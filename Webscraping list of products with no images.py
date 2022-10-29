#!/usr/bin/env python
# coding: utf-8

# In[1]:


from urllib.request import Request,urlopen
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np

prod_name=[]
img_links=[]
prod_links=[]

for pages in range(1,13):
    link='https://yoshops.com/products?page={}'
    req=Request(link.format(pages),headers={"User-Agent" :"Mozilla/5.0"})
    soup = BeautifulSoup(urlopen(req).read(),"html.parser")
    
    for item in soup.find_all("div",attrs={"class":"product"}): 
        img_links.append(item.find("div",attrs={"class":"product-thumb-inner"}).find("img").get("src"))
        prod_name.append(item.find("div",attrs={"class":"product-thumb-inner"}).find("img").get("alt"))
        prod_links.append("https://yoshops.com" + item.find("a").get("href"))

df=pd.DataFrame({"Product names":prod_name,"Product links":prod_links,"src url":img_links})
df = df[df['src url'].str.contains('noimage') == True].drop(['src url'],axis=1)
df.index = np.arange(1, len(df) + 1)
df.to_excel('C:\\Users\\Murtuza pipulyawala\\Desktop\\project_task3.xlsx')
df


# In[ ]:




