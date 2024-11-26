import openpyxl
import requests
from bs4 import BeautifulSoup
from zipfile import ZipFile
import os

# Getting Paginations Links
def get_pageinations_links(soup):
    # Getting Pagination for number of pages should iterate
    pages_iter = soup.find("ul", class_ = 'pagination')
    try:
        pages_iter_li = pages_iter.find_all("li")
        # print("Number of pages:", len(pages_iter_li))

        # Extract href attributes from <a> tags inside each <li>
        href_pages = [li.find('a')['href'] for li in pages_iter_li if li.find('a')]
        # print("Pagination hrefs:", href_pages)
        return href_pages
    except Exception as e:
        print("No pagination found", e)
        return []
    
    
# Finding all the heading of tables
def get_table_headers(soup):
    headers_text = soup.find_all('th')
    main_table_headers = [title.text.strip() for title in headers_text]
    # print(main_table_headers) # ['Team Name', 'Year', 'Wins', 'Losses', 'OT Losses', 'Win %', 'Goals For (GF)', 'Goals Against (GA)', '+ / -']
    return main_table_headers

# Function to fetch HTML content from the given URL
def fetch_html(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.text
    else:
        print(f"Failed to fetch URL: {url}, Status Code: {response.status_code}")
        return None
    
def save_html_pages(soups, page_count, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    filenaem = f"{page_count}.html" 
    file_path = os.path.join(output_dir, filenaem)
    with open(file_path, "w", encoding= 'utf-8') as html:
        html.write(str(soups.prettify()))

def save_data_sheet(soups, data):
    cell_datas = soups.find_all('tr', class_='team')
    for row_data in cell_datas:
        try:
            name = row_data.find("td", class_="name").text.strip()
            year = row_data.find("td", class_="year").text.strip()
            wins = row_data.find("td", class_="wins").text.strip()
            losses = row_data.find("td", class_="losses").text.strip()
            ot_losses = row_data.find("td", class_="ot-losses").text.strip()
            pct = row_data.find("td", class_="pct").text.strip()
            gf = row_data.find("td", class_="gf").text.strip()
            ga = row_data.find("td", class_="ga").text.strip()
            diff = row_data.find("td", class_="diff").text.strip()
            data.append((name, year, wins, losses, ot_losses, pct, gf, ga, diff))
        except AttributeError as e:
            print("Error processing row:", e)


def zip_html_pages(zip_dir, output_dir):
    with ZipFile(zip_dir, 'w') as ziper:
        for root, _, files in os.walk(output_dir):
            for file in files:
                ziper.write(os.path.join(root, file), arcname=file)


def write_to_excel(data_dicts, main_table_headers, file_name):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    # Write headers
    sheet.append(main_table_headers)
    # Write rows in sheet
    for data_dict in data_dicts:  # Assuming data_dicts is a list of dictionaries
        sheet.append([data_dict.get(header, "") for header in main_table_headers])
    # Save the workbook
    workbook.save("output.xlsx")


def calculate_winner_loser(file_name, output_sheet_name):
    # Calculating the the Top wins and loss in a year
    worksheet_name =  file_name
    rb = openpyxl.load_workbook(worksheet_name)
    sheet = rb.active
    ws = rb.create_sheet('Winner and Loser per Year')
    # Create headers for the new sheet
    ws.append(['Year', 'Winner', 'Winner Num. of Wins', 'Loser', 'Loser Num. of Wins'])
    rb = openpyxl.load_workbook(worksheet_name)
    ws = rb.create_sheet(output_sheet_name)
    # Create headers for the new sheet
    ws.append(['Year', 'Winner', 'Winner Num. of Wins', 'Loser', 'Loser Num. of Wins'])
    temp = []
    for row in sheet:
        year = row[1].value
        name = row[0].value
        wins = row[2].value
        # print(year, name, wins)
        temp.append((year,name,wins))
    result = sorted(temp, key = lambda x : x[2], reverse=False)
    hst_len = len(result)
    # Appending New Values to sheet 2 
    ws.append([result[2][1], result[2][0], result[2][2], result[hst_len - 3][1], result[hst_len - 3][0], result[hst_len - 3][2]])
    rb.save(worksheet_name)
    print(f"Sheet '{output_sheet_name}' updated successfully!")


def main():
    # Main URL to fetch
    url = "https://www.scrapethissite.com/pages/forms/"
    response = requests.get(url)
    # print(response) #<Response [200]>
    soup = BeautifulSoup(response.text, "html.parser")
    
    main_table_headers = get_table_headers(soup)
    pagination_links = get_pageinations_links(soup)
    
    zip_dir = "Collection_html.zip"
    output_dir = "HTML PAGES"
    file_name = "output.xlsx"
    page_count = 1
    data = []
    
    
    for page in pagination_links:
        page_url = f"https://www.scrapethissite.com/{page}"
        soups = BeautifulSoup(fetch_html(page_url), 'html.parser')
        save_html_pages(soups, page_count, output_dir)
        save_data_sheet(soups, data)
        page_count += 1

    zip_html_pages(zip_dir, output_dir)
    
    # Create dictionary list for writing to Excel
    data_dicts = [dict(zip(main_table_headers, row)) for row in data]
    
    write_to_excel(data_dicts, main_table_headers, file_name)
    
    # Calculate and append winners/losers per year
    calculate_winner_loser(file_name, "Winner and Loser per Year")
    
if __name__ == "__main__":
    main()