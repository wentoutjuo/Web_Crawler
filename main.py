import site
import os.path
lib_folder = os.path.join(os.path.dirname(__file__), 'lib')
site.addsitedir(lib_folder)
import requests
import xlrd
import xlwt
import logging
import shutil
from BeautifulSoup import BeautifulSoup
from time import sleep


local_proxy = {"http": "http://127.0.0.1:8087"}

colunm_type = ["product_id", "title_name", "release_date", "video_length", "actress", "maker", "label", "keyword", "description"]

target_filename = 'target.xlsx'

def QueryData(product_id):
    data = dict()
    data['product_id'] = product_id
    url = 'http://www.dmm.co.jp/search/=/searchstr=%s/analyze=V1EBDVcGUgM_/n1=FgRCTw9VBA4GCF5WXA__/n2=Aw1fVhQKX19XC15nV0AC/sort=ranking/' % product_id.replace(" ", "")
    logging.info("Querying ... %s" % url)
    try:
        searched_items = requests.get(url, proxies=local_proxy)
        logging.info("Success")
        soup = BeautifulSoup(searched_items.content)
    except requests.exceptions.ConnectionError:
        logging.info("Fail")
        return data
    finally:
        pass

    # get the target URL
    try:
        target_url = soup.find('p', {'class': 'tmb'}).find('a', href=True)['href']
    except:
        logging.info("fetch target url fail !")
        return data
    finally:
        pass

    if '/detail/=/cid' not in target_url:
        logging.info("fetch target wrong ?")
        return data

    # redirect to the target page
    target_response = requests.get(target_url, proxies=local_proxy)
    if target_response is None:
        logging.info("fetch target request fail")
        return data

    logging.info("parsing ... %s" % target_url)
    soup = BeautifulSoup(target_response.content)
    try:
        data['title_name'] = soup.find('h1', {"id": "title", "class": "item fn"}).getText()

        table = soup.find("table", {"class": "mg-b20"}).findAll("tr")
        data['release_date'] = table[1].findAll("td")[1].getText()
        data['video_length'] = table[2].findAll("td")[1].getText()
        data['actress'] = table[3].findAll("td")[1].getText()
        data['series'] = table[4].findAll("td")[1].getText()
        data['unknown'] = table[5].findAll("td")[1].getText()
        data['maker'] = table[6].findAll("td")[1].getText()
        data['label'] = table[7].findAll("td")[1].getText()
        data['keyword'] = table[8].findAll("td")[1].getText()
        clip_id = table[9].findAll("td")[1].getText()

        data['description'] = soup.find("p", {"class": "mg-b20"}).getText()

        pic_url = soup.findAll("a", {"id": clip_id}, href=True)[1]['href']
        logging.info("fetch image from %s" % pic_url)
        r = requests.get(pic_url, stream=True)
        if r.status_code == 200:
            with open('./img/%s.jpg' % product_id.replace(" ", ""), "wb") as f:
                r.raw.decode_content = True
                shutil.copyfileobj(r.raw, f)
        sleep(3)
        

    except AttributeError:
        pass
    except IndexError:
        pass
    finally:
        pass
    return data


if __name__ == '__main__':
    FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    logging.basicConfig(level=logging.INFO, format=FORMAT)
    logging.info("checking for target file: %s" % target_filename)
    if not os.path.exists('./img'):
        os.mkdir('./img')
    with xlrd.open_workbook('target.xlsx') as xlsx_input:
        try:
            table_input = xlsx_input.sheet_by_index(0)
        except IndexError, err:
            logging.critical("Page Index Error !")
            exit(0)
        finally:
            pass
        logging.info("There are %d row are waiting to be processed" % table_input.nrows)

        xlsx_output = xlwt.Workbook()
        table_output = xlsx_output.add_sheet('output')

        for i in range(table_input.nrows):
            print table_input.row_values(i)
            output = QueryData(table_input.row_values(i)[0])
            for index, item in enumerate(colunm_type):
                try:
                    table_output.write(i, index, output[item])
                except KeyError:
                    pass
            xlsx_output.save("result.xls")
