#!/usr/bin/env python
import sys, urllib2, re, xlrd, xlwt
from xlutils.copy import copy
from bs4 import BeautifulSoup as bs


def price(pn):
    url = 'http://www.digikey.com/product-detail/en/' + pn.rstrip("-ND") + "/" + pn
    part_price = ""
    try:
        content = post(pn, url)
        content.prettify()
        if content.find(id="pricing") != None:
            part_price = content.find(id="pricing").findAll("td")

            if part_price[0].text != "Call":
                part_price = part_price[1].text
            elif content.find("td", "lnkSugSub") != None:
                part_price = filt(content, "lnkSugSub")
            elif content.find("td", "lnkRohsSub") != None:
                part_price = filt(content, "lnkRohsSub")
            elif content.find("td", "lnkAltPack") != None:
                part_price = filt(content, "lnkAltPack")
            else:
                part_price = "Verify P/N"
    except urllib2.HTTPError, e:
        if e.code == 404:
            url = 'http://www.digikey.com/product-search/en?KeyWords=' + pn
            content = post(pn, url)
            pn = content.find("td", "digikey-partnumber").text.rstrip()
            if "/" in pn:
                pn = re.sub('/', '%2F', pn)
            url = 'http://www.digikey.com/product-detail/en/' + pn.rstrip("-ND") + "/" + pn
            content = post(pn, url)
            part_price = content.find(id="pricing").findAll("td")[1].text
    return part_price


def read_BOM():
    # Reading
    rb = xlrd.open_workbook('ERT.xls', formatting_info=True)
    rs = rb.sheet_by_name('Parts')
    num_rows = rs.nrows - 1
    headers = [rs.cell_value(2, i) for i in range(rs.ncols)]
    current_row = 2
    # Writing
    wb = copy(rb)
    ws = wb.get_sheet(0)
    ws.col(10).width = 256 * 25
    ws.write(current_row, 10, "Substitute P/N")
    while current_row < num_rows:
        current_row += 1
        row = rs.row(current_row)
        pn = rs.cell_value(current_row, 6)[8:]
        if rs.cell_type(current_row, 6) != 0 and ("digikey" in rs.cell_value(current_row, 6).lower()):
            if "/" in pn:
                pn = re.sub('/', '%2F', pn)
            if type(price(pn)) is tuple:
                ws.write(current_row, 7, price(pn)[0])
                ws.write(current_row, 10, price(pn)[1])
                print pn, price(pn)
            else:
                ws.write(current_row, 7, price(pn))
                print pn, price(pn)
    wb.save('ERT2.xls')


def post(pn, url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    req = urllib2.Request(url, None, headers)
    content = bs(urllib2.urlopen(req).read())
    return content


def filt(content, link):
    pn = content.find("td", link).text
    url = 'http://www.digikey.com/product-detail/en/' + pn.rstrip("-ND") + "/" + pn
    content = post(pn, url)
    if content.find(id="pricing") is None:
        return "Call"
    return content.find(id="pricing").findAll("td")[1].text, pn


def main():
    read_BOM()


main()
