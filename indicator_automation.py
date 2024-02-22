#!/usr/bin/env python
# coding: utf-8

# ### Step 1 - Import Files and Libraries

# In[44]:


#import library
import pandas as pd
#libraries for email
import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
# library for file handling
import pathlib


# In[45]:


#import and clean databases

emails = pd.read_excel(r'data_base/emails.xlsx')
stores = pd.read_csv(r'data_base/store.csv', encoding='utf-8', sep=';')
sales = pd.read_excel(r'data_base/sales.xlsx')
display(emails)
display(stores)
display(sales)


# ### Step 2 - Define Create a Table for each Store and Define the Indicator day

# In[46]:


#include store name in sales

sales = sales.merge(stores, on='ID Store')
display(sales)


# In[47]:


#create dictionary for store the stores
# loc to filter what I want
dictionary_store = {}
for store in stores['Store']:
    dictionary_store[store] = sales.loc[sales['Store']==store, :]
display(dictionary_store['Norte Shopping'])


# In[48]:


#difine the indicator day
indicator_day = sales['Date'].max()
print(indicator_day)
print('{}/{}'.format(indicator_day.day, indicator_day.month))


# ### Step 3 - Save the spreadsheet in the backup folder

# In[49]:


#identify if the folder already exists
backup_path = pathlib.Path(r'backup_stores')

files_backup_path = backup_path.iterdir()

list_names_backup = [file.name for file in files_backup_path]
#for file in files_backup_path:
#    list_names_backup.append(file.name)

#save to folder
for store in dictionary_store:
    if store not in list_names_backup:
        # to creat a path u need to pass the adress with mkdir
        new_path = backup_path / store
        new_path.mkdir()

    #config file_name to save in excel format
    file_name = '{}_{}_{}.xlsx'.format(indicator_day.month, indicator_day.day, store)
    local_path = backup_path / store / file_name

    #creat before 
    dictionary_store[store].to_excel(local_path)


# ### Step 4 - Calculate the indicator for 1 store and send by email to manager

# In[50]:


#define goals

goal_inv_day = 1000
goal_inv_year = 1650000
goal_div_prod_day = 4
goal_div_prod_year = 120
goal_averg_ticket_day = 500 
goal_averg_ticket_year = 500


# In[51]:


#goals setting

for store in  dictionary_store:
    
    store_sales = dictionary_store[store]
    store_sales_by_day = store_sales.loc[store_sales['Date'] == indicator_day, :]

    #invoicing(sum of final value)
    invoicing_year = store_sales['Final Value'].sum()
    #print(invoicing_year)
    invoicing_day = store_sales_by_day['Final Value'].sum()
    #print(invoicing_day)

    #diversity of products
    #cut all duplicate value with unique
    qtd_products_year = len(store_sales['Product'].unique())
    #print(qtd_products_year)

    qtd_products_day = len(store_sales_by_day['Product'].unique())
    #print(qtd_products_day)

    #average ticket
    value_by_sales = store_sales.groupby('Sales Code').sum('ID Store')
    average_ticket_year = value_by_sales['Final Value'].mean()
    #print(average_ticket_year)

    value_by_sale_day = store_sales_by_day.groupby('Sales Code').sum('ID Store')
    average_ticket_day = value_by_sale_day['Final Value'].mean()
    #print(average_ticket_day)
    #display(sales_store)
    #average_ticke_year =
    #average_ticke_day =

    #send email
    def send_email():  

        name = emails.loc[emails['Store']== store,'Manager'].values[0]
        msg = MIMEMultipart()
        msg['Subject'] = f'One Page Day {indicator_day.day}/{indicator_day.month} - Store {store}'
        msg['From'] = 'agathabarros@gmail.com'
        msg['To'] = emails.loc[emails['Store']== store,'E-mail'].values[0]
        password = 'sgyc osyt ajnw uywp' 


        # config email
        if invoicing_day >= goal_inv_day:
            color_inv_day = 'green'
        else:
            color_inv_day = 'red'
        if invoicing_year >= goal_inv_year:
            color_inv_year = 'green'
        else:
            color_inv_year = 'red'
        if qtd_products_day >= goal_div_prod_day:
            color_qtd_day = 'green'
        else:
            color_qtd_day = 'red'
        if qtd_products_year >= goal_div_prod_year :
            color_qtd_year = 'green'
        else:
            color_qtd_year = 'red'
        if average_ticket_day >= goal_averg_ticket_day:
            color_ticket_day = 'green' 
        else:
            color_ticket_day = 'red'
        if average_ticket_year >= goal_averg_ticket_year:
            color_ticket_year = 'green'
        else:
            color_ticket_year = 'red'


        #mail body HTML
        mail_body = f'''
        <p>Good Morning, {name}</p>

        <p>Yesterday's result <strong>{indicator_day.day}/{indicator_day.month}</strong> From <strong>Store {store}</strong> was:</p>

        <table>
        <tr>
            <th>Indicator</th>
            <th>Day Value</th>
            <th>Goal Day</th>
            <th>Day Scenario</th>
        </tr>
        <tr>
            <td>Billing</td>
            <td style="text-align: center">R${invoicing_day:.2f}</td>
            <td style="text-align: center">R${goal_inv_day:.2f}</td>
            <td style="text-align: center"><font color="{color_inv_day}">◙</font></td>
        </tr>
        <tr>
            <td>Product Diversity</td>
            <td style="text-align: center">{qtd_products_day}</td>
            <td style="text-align: center">{goal_div_prod_day}</td>
            <td style="text-align: center"><font color="{color_qtd_day}">◙</font></td>
        </tr>
        <tr>
            <td>Average Ticket</td>
            <td style="text-align: center">R${average_ticket_day:.2f}</td>
            <td style="text-align: center">R${goal_averg_ticket_day:.2f}</td>
            <td style="text-align: center"><font color="{color_ticket_day}">◙</font></td>
        </tr>
        </table>
        <br>
        <table>
        <tr>
            <th>Indicator</th>
            <th>Value Year</th>
            <th>Target Year</th>
            <th>Year Scenario</th>
        </tr>
        <tr>
            <td>Billing</td>
            <td style="text-align: center">R${invoicing_year:.2f}</td>
            <td style="text-align: center">R${goal_inv_year:.2f}</td>
            <td style="text-align: center"><font color="{color_inv_year}">◙</font></td>
        </tr>
        <tr>
            <td>Product Diversity</td>
            <td style="text-align: center">{qtd_products_year}</td>
            <td style="text-align: center">{goal_div_prod_year}</td>
            <td style="text-align: center"><font color="{color_qtd_year}">◙</font></td>
        </tr>
        <tr>
            <td>Average Ticket</td>
            <td style="text-align: center">R${average_ticket_year:.2f}</td>
            <td style="text-align: center">R${goal_averg_ticket_year:.2f}</td>
            <td style="text-align: center"><font color="{color_ticket_year}">◙</font></td>
        </tr>
        </table>

        <p>Please find attached the spreadsheet with all the dat for more details.</p>

        <p>I'm at your disposal if you have any questions.</p>
        <p>Att., Agatha</p>  '''


        msg.attach(MIMEText(mail_body, 'html'))

        #attachment withsmtlib
        attachment = pathlib.Path.cwd() / backup_path / store / f'{indicator_day.month}_{indicator_day.day}_{store}.xlsx'
        with open(attachment, 'rb') as f:
            attach = MIMEBase("application", "octet-stream")
            attach.set_payload(f.read())

            #encode the file because the file is in bytes and we need to convert to base64
            encoders.encode_base64(attach)
            attach.add_header('Content-Disposition', f'attachment; filename= {attachment.name}')
            msg.attach(attach)

        s = smtplib.SMTP('smtp.gmail.com', 587)
        #s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('E-mail Store {} sent'.format(store))


    send_email()


# In[52]:


#ranking of stores year and day
invoicing_stores = sales.groupby('Store')[['Store', 'Final Value']].sum('Final Value')
invoicing_stores = invoicing_stores.sort_values(by='Final Value', ascending=False)
display(invoicing_stores)

#filterin by day with loc
sales_day = sales.loc[sales['Date'] == indicator_day, :]
invoincing_stores_day = sales_day.groupby('Store')[['Store', 'Final Value']].sum('Final Value')
invoincing_stores_day = invoincing_stores_day.sort_values(by='Final Value', ascending=False)
display(invoincing_stores_day)


# ### Step 7 - Create ranking for board

# In[53]:


#ranking of stores year and day
invoicing_stores = sales.groupby('Store')[['Store', 'Final Value']].sum('Final Value')
invoicing_stores_year = invoicing_stores.sort_values(by='Final Value', ascending=False)
display(invoicing_stores_year)

file_name = '{}_{}_ranking_year.xlsx'.format(indicator_day.month, indicator_day.day)
invoicing_stores_year.to_excel(r'backup_stores/{}'.format(file_name))

#filterin by day with loc
sales_day = sales.loc[sales['Date'] == indicator_day, :]

invoincing_stores_day = sales_day.groupby('Store')[['Store', 'Final Value']].sum('Final Value')
invoincing_stores_day = invoincing_stores_day.sort_values(by='Final Value', ascending=False)
display(invoincing_stores_day)

file_name = '{}_{}_ranking_day.xlsx'.format(indicator_day.month, indicator_day.day)
invoicing_stores_year.to_excel(r'backup_stores/{}'.format(file_name))


# ### Step 8 - Send email to management

# In[54]:


def send_email_board():  

        msg = MIMEMultipart()
        msg['Subject'] = f'Day Ranking {indicator_day.day}/{indicator_day.month}'
        msg['From'] = 'agathabarros@gmail.com'
        msg['To'] = emails.loc[emails['Store']== 'Board','E-mail'].values[0]
        password = 'sgyc osyt ajnw uywp' 


       

        #mail body HTML
        mail_body = f''' 
        Dears, good morning

        Best Store of the Day in Revenue: Store {invoincing_stores_day.index[0]} with Revenue R${invoincing_stores_day.iloc[0, 0]:.2f}
        Worst store of the day in Revenue: Store {invoincing_stores_day.index[-1]} with Revenue R${invoincing_stores_day.iloc[-1, 0]:.2f}

        Best Store of the Year in Revenue: Store {invoicing_stores_year.index[0]} with Revenue R${invoicing_stores_year.iloc[0, 0]:.2f}
        Worst store of the Year in Revenue: Store {invoicing_stores_year.index[-1]} with Revenue R${invoicing_stores_year.iloc[-1, 0]:.2f}

        Attached are the year and day rankings for all stores.

        Any doubt I am available.

        Att.,
        Lira

    '''

        print(mail_body)
        msg.attach(MIMEText(mail_body, 'html'))

        #attachment withsmtlib
        attachment = pathlib.Path.cwd() / backup_path / f'{indicator_day.month}_{indicator_day.day}_ranking_year.xlsx'
        with open(attachment, 'rb') as f:
            attach = MIMEBase("application", "octet-stream")
            attach.set_payload(f.read())

            #encode the file because the file is in bytes and we need to convert to base64
            encoders.encode_base64(attach)
            attach.add_header('Content-Disposition', f'attachment; filename= {attachment.name}')
            msg.attach(attach)
        
        attachment = pathlib.Path.cwd() / backup_path / f'{indicator_day.month}_{indicator_day.day}_ranking_day.xlsx'
        with open(attachment, 'rb') as f:
            attach = MIMEBase("application", "octet-stream")
            attach.set_payload(f.read())

            #encode the file because the file is in bytes and we need to convert to base64
            encoders.encode_base64(attach)
            attach.add_header('Content-Disposition', f'attachment; filename= {attachment.name}')
            msg.attach(attach)

        s = smtplib.SMTP('smtp.gmail.com', 587)
        #s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('E-mail Board sent.')

send_email_board()


# 

# In[ ]:




