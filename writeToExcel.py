import re
from xlwt import Workbook
from bs4 import BeautifulSoup
import requests
import os

hills_file = "hills.xls"
ski_central_url = "https://www.skicentral.com/"


def make_workbook():
    url_wi_hills = ski_central_url + "wisconsin.html"
    req = requests.get(url_wi_hills)
    soup_hill_list = BeautifulSoup(req.content, "html.parser")
    wb = Workbook()

    # Check for existing workbook. If exists, delete. If not, create file
    if os.path.exists(hills_file):
        os.remove(hills_file)
        # TODO: Change to return once program is done

    hill_wb = wb.add_sheet("Hill List")
    hill_wb.write(0, 0, 'NAME')
    hill_wb.write(0, 1, 'URL')
    hill_wb.write(0, 2, 'STATUS')

    hill_count = 1
    # Check if hill_list_block is empty
    try:
        hill_div_list = soup_hill_list.find_all('div', {'class': ['listitemline']})
        for hill_div in hill_div_list:
            hill_name = hill_div['data-name']
            url_str = ski_central_url + hill_div['data-link']
            hill_url = get_url_html(url_str)
            # TODO: hill_status
            hill_status = "Fix Me"
            # Write hill info to workbook
            write_to_workbook(hill_wb, hill_name, hill_url, hill_status, hill_count)
            hill_count += 1
        # Print excel file to ensure correct data

        wb.save(hills_file)
    except AttributeError:
        print('hill_list_block is empty.')


def get_url_html(url):
    req = requests.get(url)
    soup_hill = BeautifulSoup(req.content, "html.parser")
    hill_website = soup_hill.find(text=re.compile("Website:"))
    if hill_website == None:
        return "N/A"
    else:
        return hill_website.next_element.text


# Function that writes hill info to workbook
def write_to_workbook(hill_wb, hill_name, hill_url, hill_status, hill_count):
    print("HILL NAME: %s\nHILL URL: %s\nHILL STATUS: %s\n" % (hill_name, hill_url, hill_status))
    hill_wb.write(hill_count, 0, hill_name)
    hill_wb.write(hill_count, 1, hill_url)
    hill_wb.write(hill_count, 2, hill_status)
    return
