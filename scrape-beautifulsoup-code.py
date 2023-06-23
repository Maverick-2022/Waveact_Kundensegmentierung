import requests
from bs4 import BeautifulSoup
import pandas as pd

url = 'https://cryptorank.io/funding-rounds/'
response = requests.get(url)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')
    table = soup.find('table')

    if table is not None:
        headers = []
        for th in table.find_all('th'):
            header_text = th.text.strip()
            if header_text in ['Project', 'Date', 'Raise', 'Stage', 'Funds and Investors', 'Category']:
                headers.append(header_text)

        rows = []
        for tr in table.find_all('tr'):
            row = []
            for td in tr.find_all('td'):
                if td.find('a'):
                    link_text = td.find('a').text.strip()
                    row.append(link_text)
                else:
                    cell_text = td.text.strip()
                    row.append(cell_text)
            if row:
                filtered_row = [row[i] for i in range(len(headers))]
                rows.append(filtered_row)

        df = pd.DataFrame(rows, columns=headers)
        df['Funds and Investors'] = df['Funds and Investors'].apply(lambda x: ' '.join(x.split()[:-1]) if isinstance(x, str) else x)
        print(df)
        df.to_excel('D:\Refaatdata.xlsx', index=False)
        print("Data saved to data.xlsx")
    else:
        print("Table element not found on the webpage")
else:
    print("Failed to retrieve webpage")
