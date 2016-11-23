#coding:utf-8  
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import sys,os  
import urllib
import urllib2
import urllib3
import cookielib
import requests
import xlrd
import xlwt
import time
from xlrd import open_workbook
from xlutils.copy import copy
from PIL import Image,ImageDraw  
from bs4 import BeautifulSoup
def get_all(id):
    cookies = {}

    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/50.0.2661.86 Safari/537.36'
    }

    def get_code():
        url = 'http://222.195.242.222:8080/validateCodeAction.do'
        resp = requests.get(url, headers=headers)
        cookies['JSESSIONID'] = resp.cookies.get('JSESSIONID')
        time.sleep(0.010)
        with open('Z:/python/code.jpg', 'wb') as img:
            img.write(resp.content)

    def CodeRecognition():
    
        # urllib.urlretrieve('http://222.195.242.222:8080/validateCodeAction.do','D:/orginal.jpg')
        im = Image.open('Z:/python/code.jpg')
    
        # convert to grey level image转化为灰度图
        Lim = im.convert('L')
        Lim.save('d:/fun_Level.jpg')
    
        # setup a converting table with constant threshold
        threshold = 127
        table = []
        for i in range(256):
            if i < threshold:
                table.append(0)
            else:
                table.append(1)
    
        # convert to binary image by the table
        bim = Lim.point(table, '1')
    
        # bim.save('Z:/python/fun_binary.jpg')
    
    
    
          
        #二值判断,如果确认是噪声,用改点的上面一个点的灰度进行替换  
        #该函数也可以改成RGB判断的,具体看需求如何  
        def getPixel(image,x,y,G,N):  
            L = image.getpixel((x,y))  
            if L > G:  
                L = True  
            else:  
                L = False  
          
            nearDots = 0  
            if L == (image.getpixel((x - 1,y - 1)) > G):  
                nearDots += 1  
            if L == (image.getpixel((x - 1,y)) > G):  
                nearDots += 1  
            if L == (image.getpixel((x - 1,y + 1)) > G):  
                nearDots += 1  
            if L == (image.getpixel((x,y - 1)) > G):  
                nearDots += 1  
            if L == (image.getpixel((x,y + 1)) > G):  
                nearDots += 1  
            if L == (image.getpixel((x + 1,y - 1)) > G):  
                nearDots += 1  
            if L == (image.getpixel((x + 1,y)) > G):  
                nearDots += 1  
            if L == (image.getpixel((x + 1,y + 1)) > G):  
                nearDots += 1  
          
            if nearDots < N:  
                return image.getpixel((x,y-1))  
            else:  
                return None  
          
        # 降噪   
        # 根据一个点A的RGB值，与周围的8个点的RBG值比较，设定一个值N（0 <N <8），当A的RGB值与周围8个点的RGB相等数小于N时，此点为噪点   
        # G: Integer 图像二值化阀值   
        # N: Integer 降噪率 0 <N <8   
        # Z: Integer 降噪次数   
        # 输出   
        #  0：降噪成功   
        #  1：降噪失败   
        def clearNoise(image,G,N,Z):  
            draw = ImageDraw.Draw(image)  
          
            for i in xrange(0,Z):  
                for x in xrange(1,image.size[0] - 1):  
                    for y in xrange(1,image.size[1] - 1):  
                        color = getPixel(image,x,y,G,N)
                        if color != None:  
                            draw.point((x,y),color)  
    
        def main():  
            #打开图片  
            # bim = Image.open("Z:/python/fun_binary.jpg")  
          
            # #将图片转换成灰度图片  
            # bim = image.convert("L")  
          
            #去噪,G = 50,N = 4,Z = 4  
            clearNoise(bim,10,2,1)  
          
            #保存图片  
            bim.save("Z:/python/result.jpg")  
          
          
        main()
    
        def image_to_string(img, cleanup=True, plus='-7'):
            # cleanup为True则识别完成后删除生成的文本文件
            # plus参数为给tesseract的附加高级参数
            os.popen('tesseract ' + img + ' ' + img + ' ' + plus)  # 生成同名txt文件
            text = file(img + '.txt').read().strip()
            if cleanup:
                os.remove(img + '.txt')
            return text
    
        m_code = image_to_string('Z:/python/result.jpg')
        
        print m_code
        return m_code

    get_code()
    code = CodeRecognition()

    while len(code) != 4:
        get_code()
        code = CodeRecognition()

    
    def login():
        # code = input('input the code: ')
        url = 'http://222.195.242.222:8080/loginAction.do'
        form = {
            'zjh1': '',
            'tips': '',
            'lx': '',
            'evalue': '',
            'eflag': '',
            'fs': '',
            'dzslh': '',
            'zjh': str(id),
            'mm': str(id),
            'v_yzm': code
        }
        resp = requests.post(url, headers=headers, data=form, cookies=cookies)
    
    
    def get_info():
        url = 'http://222.195.242.222:8080/bxqcjcxAction.do'
        resp = requests.get(url, headers=headers, cookies=cookies)
        html = resp.text
        # 定义一个beautifulsoup对象
        soup = BeautifulSoup(html,"html.parser")
        data = soup.find_all()

        info = ['0','0','0','0','0']
        index = html.find('信号')
        if index > 1:
            
            # student_id = data[0].next_sibling.next_sibling.string
            # student_id = student_id.strip()
            # address = data[21].previous_sibling.previous_sibling.string
            # address = address.strip()
            # academy = data[24].next_sibling.next_sibling.string
            # academy = academy.strip()
            # profession = data[25].next_sibling.next_sibling.string
            # profession = profession.strip()
            # name = data[1].next_sibling.next_sibling.string
            # name = name.strip()

            
            # info[0] = student_id
            # info[1] = name
            # info[2] = academy
            # info[3] = profession
            # info[4] = address

            # print student_id
            # print name
            # print academy
            # print profession
            # print address
            print data[0].parent.next_sibling.next_sibling.next_sibling.next_sibling.string
        else:
            print str(id)
            info = [id,'0','0','0','0']
        
        return info
    login()
    info = get_info()
    return info

def fff(start,end):
    i = 0
    lists = [[0 for col in range(5)] for row in range(500)]
    for id in range(start,end):
        info = get_all(id)
        lists[i][0] = info[0]
        lists[i][1] = info[1]
        lists[i][2] = info[2]
        lists[i][3] = info[3]
        lists[i][4] = info[4]
        i += 1
    return lists

if __name__ == '__main__':
    lists = fff(201428416,201428478)

    rb = open_workbook('demo.xls')
    
    rs = rb.sheet_by_index(0)
    wb = copy(rb)
    
    #通过get_sheet()获取的sheet有write()方法
    ws = wb.get_sheet(0)

    for i in range(0,499):
        ws.write(i,0,lists[i][0])
        ws.write(i,1,lists[i][1])
        ws.write(i,2,lists[i][2])
        ws.write(i,3,lists[i][3])
        ws.write(i,4,lists[i][4])
    
    #wb.save('demo.xls')
