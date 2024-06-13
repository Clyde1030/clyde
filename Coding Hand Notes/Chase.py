from bs4 import BeautifulSoup
import pandas as pd
import xlwings as xw
import logging
import os

LOGGER = logging.getLogger(__name__)

def main():
    set_logger('info')
    wb_path = '/Users/yu-shenglee/Library/CloudStorage/OneDrive-Personal/Desktop/Excel on Desktop/dev/Python Projects/Stock Portfolio-M.xlsx'
    
    names = ['Clyde', 'Austin', 'Irina']
    
    for name in names:
        with open('html_' + name + '.txt', 'r') as file:
            html_long = file.readlines()
        df = extract(html_long[0])
        LOGGER.info(df)
        output(df, wb_path, name)

    # with open('html_Austin.txt','r') as file:
    #     html_long = file.readlines()
    # df = extract(html_long[0])
    # LOGGER.info(df)
    # output(df, wb_path, 'Austin')

    # with open('html_Irina.txt','r') as file:
    #     html_long = file.readlines()
    # df = extract(html_long[0])
    # LOGGER.info(df)
    # output(df, wb_path, 'Irina')

def output(df, path, who):
    LOGGER.debug(f'Accessing Excel workbook {os.path.basename(path)}...')
    if df is not None:
        data = df.values.tolist()
        with xw.App(visible=True) as Excel:
            xw.App.display_alerts = False
            wb = xw.Book(path)
            ws = wb.sheets['Pandas']
            from_point = ws.range(who)
            to_point = ws.range(who).offset(100,2)
            LOGGER.debug(f'Clear content from {str(from_point)} to {str(to_point)}')
            ws.range(from_point, to_point).clear_contents()            
            ws.range(who).value = data
            wb.save()
    LOGGER.info(f'Finish pasting data for {who}.')

def extract(html):
    '''take the HTML text as input and return the pandas dataframe as output'''
    html_long_data = BeautifulSoup(html, 'html.parser') 
    tables = html_long_data.find_all('tbody')
    rows = tables[0].find_all('tr')
    df = pd.DataFrame(columns=['ticker','quantity','cost'])
    for row in rows:
        cols = row.find_all('td') 
        if len(cols)!=0 and cols[2] is not None and cols[3] is not None:
            ticker = cols[0].find('mds-link').get('text')
            quantity = cols[2].string
            cost = cols[3].contents[0].string
            data_dict = {'ticker': ticker,
                        'quantity': quantity, 
                        'cost': cost}
            df1 = pd.DataFrame(data_dict, index=[0])
            df = pd.concat([df, df1], ignore_index=True)
    # df['cost'] = df['cost'].replace(',', '')    
    # df['quantity'] = df['quantity'].astype('float')
    df.loc[df['ticker']=='Cash &amp; Sweep Funds','quantity'] = 1
    df.loc[df['ticker']=='Cash &amp; Sweep Funds','ticker'] = 'Cash'
    return df

def set_logger(default_level = 'info'):
    LOGGER = logging.getLogger()
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    LOGGER.addHandler(handler)
    log_level = default_level.upper()
    LOGGER.setLevel(log_level)

if __name__=='__main__':
    main()
