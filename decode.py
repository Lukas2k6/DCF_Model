import requests
from bs4 import BeautifulSoup

def decode_secret_message(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')
    rows = soup.find_all('tr')
    grid = {}
    max_x = 0
    max_y = 0
    for row in rows[1:]:
        cols = row.find_all('td')
        if len(cols) == 3:
            x_text = cols[0].get_text().strip()
            char = cols[1].get_text().strip()
            y_text = cols[2].get_text().strip()
            if x_text.lstrip('-').isdigit() and y_text.lstrip('-').isdigit():
                x = int(x_text)
                y = int(y_text)
                grid[(x, y)] = char
                max_x = max(max_x, x)
                max_y = max(max_y, y)
    for y in range(max_y, -1, -1):
        row_str = ""
        for x in range(max_x + 1):
            row_str += grid.get((x, y), " ")
        print(row_str)

decode_secret_message("https://docs.google.com/document/d/e/2PACX-1vSvM5gDlNvt7npYHhp_XfsJvuntUhq184By5xO_pA4b_gCWeXb6dM6ZxwN8rE6S4ghUsCj2VKR21oEP/pub")