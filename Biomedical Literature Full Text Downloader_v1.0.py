#!/usr/bin/env python
# -*- coding: utf-8 -*-


from tkinter import *
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter.scrolledtext import ScrolledText

from selenium import webdriver
from lxml import etree

import time
import os

import re

import requests
import random



class DOC_GUI():
    def __init__(self,init_window_name):
        self.init_window_name = init_window_name
        self.record = {"PMID":[],
                       "title":[],
                       "journal":[],
                       "doi":[],
                       "是否成功下载":[]}  
        self.unique_data_y = []

        
        


    #设置窗口
    def set_init_window(self):
        self.init_window_name.title("Biomedical Literature Full Text Downloader_v1.0")           #窗口名
        #799 441为窗口大小，+10 +10 定义窗口弹出时的默认展示位置
        self.init_window_name.geometry('799x510+10+10')
        #self.init_window_name["bg"] = "pink"                                    
        #窗口背景色，其他背景色见：blog.csdn.net/chl0000/article/details/7657887
        
        #self.init_window_name.attributes("-alpha",0.9)                          
        #虚化，值越小虚化程度越高
        
        #标签
        self.init_data_label = Label(self.init_window_name,
                                     text="输入PMID文件",
                                     width = 100,
                                     height = 20)
        self.init_data_label.place(x = 29,y = 16,
                                   width = 100,
                                   height = 20)
        
        self.save_data_label = Label(self.init_window_name,
                                     text="输入保存文件夹",
                                     width = 100,
                                     height = 20)
        self.save_data_label.place(x = 29,
                                   y = 66,
                                   width = 100,
                                   height = 20)
        
        
        #self.init_data_label.grid(row=0, column=0)
        self.result_data_label = Label(self.init_window_name, 
                                       text="输出结果:",
                                       width = 10,
                                       height = 4)
        self.result_data_label.place(x = 19,y = 136,
                                     width = 100,height = 20)
        

        #文本框
        
        #输入路径
        self.path_inp = StringVar() 
        #保存路径
        self.path_out = StringVar() 
        
        #输入PMID路径
        self.e1 = Entry(self.init_window_name,
                        textvariable = self.path_inp,
                        width=67)
        self.e1.place(x = 162,y = 12,
                      width = 434,height = 33)
        
        #保存路径
        self.e2 = Entry(self.init_window_name,
                        textvariable = self.path_out,
                        width=67)
        self.e2.place(x = 162,
                      y = 59,
                      width = 434,
                      height = 33)   #保存路径
       
    
        #结果框
        self.result_data_Text = ScrolledText(self.init_window_name)  #处理结果展示
        self.result_data_Text.place(x = 38,
                                    y = 159,
                                    width = 557,
                                    height = 310)

        self.result_data_Text.insert(END,"开始下载时程序显示未响应,实际上在运行\
                                        \n如果长时间不刷新可以点击:“详细下载结果”按钮,结束下载并查看结果\
                                        \n如果遇到网络问题停止下载，可以点击“开始”按钮重新下载，不会重复下载\
                                        \n推荐chromedriver下载地址：\
                                        \nhttp://chromedriver.storage.googleapis.com/index.html\
                                        \nhttps://npm.taobao.org/mirrors/chromedriver/ \
                                        \n")
        
        
        
        #check框
        #默认为0，即不获取标题等信息     
        self.checkVar = StringVar(value="0")
        
        self.CheckButton_16 = Checkbutton(self.init_window_name,text="获取标题等信息",
                                          variable=self.checkVar)
        self.CheckButton_16.place(x = 635,
                                  y = 116,
                                  width = 100,
                                  height = 20)
        
        #按钮
        # PMID文件路径
        self.get_button = Button(self.init_window_name, 
                                 text="获取路径", 
                                 width=10,
                                 height = 4,
                                 command=self.selectPath_input)  
        self.get_button.place(x = 650,
                              y = 12,
                              width = 100,
                              height = 28)
        
        # 保存路径
        self.save_button = Button(self.init_window_name, text="获取路径", 
                                  width=10,
                                  height = 4,
                                  command=self.selectPath_output)  
        self.save_button.place(x = 650,
                               y = 63,
                               width = 100,
                               height = 28)
        
        # 开始
        self.start_button = Button(self.init_window_name,
                                   text="开始", 
                                   width=10,
                                   height = 4,
                                   command=self.show)  
        self.start_button.place(x = 628,
                                y = 168,
                                width = 121,
                                height = 58)
        
        #保存文件夹，默认为工作目录+test
        self.filefold_button = Button(self.init_window_name, 
                                      text="文献保存文件夹", 
                                      width=10,
                                      height = 4,
                                      command=self.opendir)  
        self.filefold_button.place(x = 628,
                                   y = 260,
                                   width = 121,
                                   height = 54)
        
        #详细下载结果按钮  到result.csv
        self.result_button = Button(self.init_window_name, 
                                    text="详细下载结果", 
                                    width=10,
                                    height = 4,
                                    command=self.getfile)  
        self.result_button.place(x = 628,
                                 y = 337,
                                 width = 120,
                                 height = 53)     
        
        #清空输出框 
        self.clear_button = Button(self.init_window_name, 
                                    text="清空输出结果", 
                                    width=10,
                                    height = 4,
                                    command=self.clearit)  
        self.clear_button.place(x = 628,
                                 y = 407,
                                 width = 120,
                                 height = 53)  
        
    
    
        
    def show(self):
        #print("开始下载:")
        
        self.result_data_Text.insert(END,"开始下载：\n")
        self.init_window_name.update()
        #self.result_data_Text.update_idletasks() 
        
        if self.checkVar.get() == "0":
            if len(self.e2.get())==0:
                self.doc_load(filename=self.e1.get())
            else:    
                self.doc_load(filename=self.e1.get(),
                          savepath=self.e2.get())
        else:
            if len(self.e2.get())==0:
                self.doc_load(filename=self.e1.get(),
                              get_titles="T"
                              )
            else:
                self.doc_load(filename=self.e1.get(),
                              savepath=self.e2.get(),
                              get_titles="T")    
            
    #功能函数
    def selectPath_input(self):
        pathinp_ = askopenfilename() #选择文件
        self.path_inp.set(pathinp_)
    
    def selectPath_output(self):
        pathout_ = askdirectory()   # 选择目录
        self.path_out.set(pathout_)    
        
        
    def opendir(self):
        #打开保存文件夹
        if os.path.exists("articles") == True or os.path.exists(self.path_out.get()) == True:
            if len(self.path_out.get()) == 0:
                os.system('start '+ "articles")
            else:
                #os.system("start "+self.path_out.get())
                self.result_data_Text.insert(END,self.path_out.get()+"------啥情况？\n")
                os.startfile(self.path_out.get())
                
        else:
            self.result_data_Text.insert(END,"结果文件夹不存在\n")

    def getfile(self):
        #若开了浏览器，则关掉
        if "self.browser" in locals().keys():
            self.browser.quit()

        #打开结果文件：result.txt
        if len(self.unique_data_y) == 0:
            self.result_data_Text.insert(END,"结果文件result.txt不存在\n")
            return


        self.result_data_Text.insert(END,"不重复PMID："+ str(len(self.unique_data_y)) +"\n" +
                                         "目标文件夹下成功获得：" + str(len(self.real_success)) + "\n" +
                                         "未成功下载：" + str(len(self.no_data))+ "\n" +
                                         "详细情况请查看工作目录下的result.txt文件\n")
        self.init_window_name.update()        
        
        file = open('result.txt', 'w',encoding = "utf-8") 
        #将字典写入
        for i in self.record.keys():
            file.write(str(i)+"\t")
        file.write("\n")
        PMID_sum = len(list(self.record.values())[0])
        
        if PMID_sum > 0:
            for i in range(0,PMID_sum):
                for v in list(self.record.values()):
                    file.write(str(v[i])+"\t")
                file.write("\n")
                # 注意关闭文件
        file.close()


        if os.path.exists("result.txt") == True:   
            os.system('start '+ "result.txt")
        else:
            self.result_data_Text.insert(END,"结果文件result.txt不存在\n")
        
        
    def clearit(self):
        self.result_data_Text.delete('1.0','end')

            
    #文献下载代码
    def download(self,PMID,requests_pdf_url,savepath):
        header = {
            "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) \
            AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36"
            }
        
        if os.path.exists(savepath) == False:
            os.mkdir(savepath)
            
        filename = savepath+"\\"+str(PMID)+ ".pdf"
        if os.path.exists(filename) == False:
            try:    
                r = requests.get(requests_pdf_url,
                                 stream=True,
                                 headers=header)
            except:
                print("下载失败" + PMID)
                self.result_data_Text.insert(END,"下载失败" + PMID + "\n")
                self.init_window_name.update()
                return
            with open(filename, 'wb+') as f:
                f.write(r.content)
            #print("成功下载："+PMID) 
            self.result_data_Text.insert(END, "成功下载：" + PMID + "\n")
            self.init_window_name.update()
        else:
            #print("目标文件夹下已经存在：" + PMID)
            self.result_data_Text.insert(END, "目标文件夹下已经存在：" + PMID + "\n")
    
    
    def pubmed_info(self,PMID):
        if self.get_titles == "T" :
        
            self.browser.set_page_load_timeout(20)
           
            try:
                self.browser.get("https://pubmed.ncbi.nlm.nih.gov/"+str(PMID))        
                
            except:
                self.browser.execute_script('window.stop()')
                
                self.browser.quit()
                self.browser = webdriver.Chrome(executable_path=r"chromedriver.exe",
                                                options=self.option)
                self.record["PMID"].append(PMID)
                self.record["title"].append("略")
                self.record["journal"].append("略")
                self.record["doi"].append("略")
                self.record['是否成功下载'].append("是")
                return self.record
                
            rpage = self.browser.page_source
                    
            tree = etree.HTML(rpage)
            
            
          #  jour = tree.xpath('//*[@id="full-view-journal-trigger"]/text()')
            jour = tree.xpath('/html/head/meta[@name="citation_publisher"]/@content')[0]
            doi = tree.xpath('/html/head/meta[@name="citation_doi"]/@content')[0]
            title = tree.xpath('/html/head/meta[@name="description"]/@content')[0]
        
            self.record["PMID"].append(PMID)
            self.record["title"].append(title)
            
            self.record["journal"].append(jour)
            self.record["doi"].append(doi)
            self.record['是否成功下载'].append("是")
            
        else:
            self.record["PMID"].append(PMID)
            self.record["title"].append("略")
            self.record["journal"].append("略")
            self.record["doi"].append("略")
            self.record['是否成功下载'].append("是")
        return self.record
    
    
    
    
    def doc_load(self,filename,savepath="articles",get_titles = "N"):
        self.get_titles = get_titles
        self.unique_data_y = []
        #读取csv文件
        """if filename.split(".")[-1] == "csv":
            
            data_x= pd.read_csv(filepath_or_buffer = filename, sep = ',').values
            data_y = list(data_x)
            data_y = data_x.tolist()
            
            unique_data_y = []
            for k in data_y:
                unique_data_y.append(str(k[0]))"""
                
        #读取txt文件
        if filename.split(".")[-1] == "txt":
             f = open(filename,"r")   #设置文件对象
             data_y = f.read().splitlines() #直接将文件中按行读到list里
             f.close()             #关闭文件
             
             self.unique_data_y = []
             for k in data_y:
                 if len(k) > 0:
                     self.unique_data_y.append(k.strip())
        else:
            #print("请输入txt或csv格式,一行一个PMID,无需行名或者列名")
            self.result_data_Text.insert(END ,"请输入txt文本文件,一行一个PMID,无需行名或者列名\n")
            self.result_data_Text.update_idletasks() 
            return
            
        self.unique_data_y = sorted(list(set(self.unique_data_y)))
        
        
    
    
                
        
        self.option = webdriver.ChromeOptions()
        self.option.add_argument('--no-sandbox')
        self.option.add_argument('--disable-gpu')
        #self.option.add_argument('--disable-software-rasterizer')
        
        
        #浏览器要求接受网站的证书。可以设置默认情况下忽略这些错误，以免发生这些错误。
        self.option.add_argument('-ignore-certificate-errors')
        self.option.add_argument('-ignore -ssl-errors')
        
        
        #self.option.add_argument('--log-level=3')
        #self.option.add_experimental_option('excludeSwitches', ['enable-logging'])
        #self.option.add_experimental_option('excludeSwitches', ['enable-automation'])
        
        self.option.add_argument('headless')  #不打开浏览器窗口
        
        
        #不加载图片
        self.option.add_argument('blink-settings=imagesEnabled=false') 
        
        #self.option.add_argument('incognito')  #隐身模式
        #self.option.add_argument('no-startup-window') 
        
        
            # 实例化一个chrome浏览器对象
        #get直接返回，不再等待界面加载完成
       # desired_capabilities = DesiredCapabilities.CHROME
       # desired_capabilities["pageLoadStrategy"] = "none"
       
        #ChromeDriverService service = ChromeDriverService.CreateDefaultService(System.AppDomain.CurrentDomain.BaseDirectory.ToString());
        #service.HideCommandPromptWindow = true;
        #webDriver = new ChromeDriver(service, option);
        try:
            
        
            self.browser = webdriver.Chrome(executable_path=r"chromedriver.exe", 
                                            options=self.option)
            
        except:
            self.result_data_Text.insert(END ,"请在程序exe旁边放入浏览器版本对应的chromedriver \n")
            
        false_data = []
        
    
        
      #使用sci-hub网站进行下载
        for i in self.unique_data_y:
            #i = "33080015"
            
            #data_y.index([29054094])
            
            PMID = i
            
            pdfname = savepath+"\\"+str(PMID)+ ".pdf"
            if os.path.exists(pdfname) == False:
                time.sleep(random.randint(2,3))   
                self.browser.get("https://sci-hub.se")        
                try:
                    elem = self.browser.find_element_by_name ("request")
                except:
                    time.sleep(60)
                    
                #    self.browser.quit()
                    
                #    self.browser = webdriver.Chrome(chrome_options=option)
                    self.browser.get("https://sci-hub.st") 
                    try:
                        elem = self.browser.find_element_by_name ("request")
                    except:
                        time.sleep(100)
                     #   self.browser.quit()
                      #  self.browser = webdriver.Chrome(chrome_options=option)
                        self.browser.get("https://sci-hub.ru") 
                        elem = self.browser.find_element_by_name ("request")
            else:
                #print("目标文件夹下已存在：" + PMID)
                self.result_data_Text.insert(END,"目标文件夹下已存在：" + PMID + "\n")
                self.init_window_name.update()
                continue

            elem.send_keys(PMID)    
            
            self.browser.find_element_by_id("open").click()
                
            rpage = self.browser.page_source
            
            lianjie = re.findall("https:.*.pdf", rpage)
            lianjie_2 = re.findall("sci-hub.*.pdf", rpage)
            
            if len(lianjie)<1 and len(lianjie_2) < 1:
                time.sleep(2)
                false_data.append(PMID)
                continue
            
            else:

                self.record = self.pubmed_info(PMID)

                if len(lianjie) >=1:
                    requests_pdf_url = lianjie[0]
                elif len(lianjie_2) >=1:
                    requests_pdf_url = lianjie_2[0]
                    requests_pdf_url =   "http://" + requests_pdf_url

                
                
                self.download(PMID,requests_pdf_url,savepath)
    
                        
        #使用pubmed进行下载
        for k in false_data:
            
        #        k = "31785987"
            PMID = k
            
            self.browser.set_page_load_timeout(30)
            time.sleep(random.randint(2,3))
            pdfname = savepath+"\\"+str(PMID)+ ".pdf"
            if os.path.exists(pdfname) == False:
                try:
                    self.browser.get("https://pubmed.ncbi.nlm.nih.gov/"+str(PMID))        
                except:
                    self.browser.execute_script('window.stop()')
                    self.browser.quit()
                    self.browser = webdriver.Chrome(executable_path=r"chromedriver.exe",
                                                    chrome_options=self.option)
                    
                    continue
            else:
                
                continue
            #self.browser.get("https://www.ncbi.nlm.nih.gov/pmc/articles/PMC4680451/")
            
            rpage = self.browser.page_source
            
            #if "free-label" in rpage:
            if 1 ==1:
                tree = etree.HTML(rpage)
                lianjie = tree.xpath('//*[@id="article-page"]/aside/div/div[1]/div/div/a/@href')
             #   //*[@id="article-page"]/aside/div/div[1]/div[1]/div/a
                
                #要是有直接的pdf链接的话：
                if len(re.findall("http.*?[.]pdf", rpage))>0:
                    self.record = self.pubmed_info(PMID)
                    requests_pdf_url = lianjie[0]
                    self.download(PMID,requests_pdf_url,savepath)
                
                #再转一层
                elif len(lianjie) > 0:
                    for j in lianjie:
                        #j = lianjie[0]
                        time.sleep(random.randint(3,4))
                       # self.browser.get(j)
                        self.browser.set_page_load_timeout(20)
                        
                        pdfname = savepath+"\\"+str(PMID)+ ".pdf"
                        if os.path.exists(pdfname) == False:
                            try:
                                self.browser.get(j)        
                            except:
                                self.browser.execute_script('window.stop()')
                                self.browser.quit()
                                self.browser = webdriver.Chrome(executable_path=r"chromedriver.exe",
                                                                chrome_options=self.option)
                                continue
                        else:
                            continue
                        pdfpage = self.browser.page_source 
                        
                        tree_pdf = etree.HTML(pdfpage)
                        
                        #通过PMC数据库下载
                        pdf_lianjie = re.findall(r"/pmc.*.pdf", pdfpage)
                        
                        if len(pdf_lianjie) >= 1: #通过PMC数据库下载
                            self.record = self.pubmed_info(PMID)
                            requests_pdf_url = pdf_lianjie[0]
                            requests_pdf_url = "https://www.ncbi.nlm.nih.gov"+requests_pdf_url
                            #self.record = self.pubmed_info(tree,PMID)
                            self.download(PMID,requests_pdf_url,savepath)
                            
                        #通过 BPG网站下载    
                        elif "全文 (PDF)" in pdfpage:
                            pdf_lianjie = tree_pdf.xpath('/html/body/div/div[1]/div[4]/div[1]/div[2]/ul/li[3]/a/@href')
                            
                            if len(pdf_lianjie) > 0:
                                requests_pdf_url = pdf_lianjie[0]
                                self.record = self.pubmed_info(PMID)
                                self.download(PMID,requests_pdf_url,savepath)
                                
                        
                        elif len(re.findall(r"/article.*[.]pdf", pdfpage)) > 0:
                            #通过journal.waocp网站
                            pdf_lianjie = re.findall(r"/article.*[.]pdf", pdfpage)
                            
                            requests_pdf_url = "http://journal.waocp.org/"+pdf_lianjie[0]
                            
                            self.record = self.pubmed_info(PMID)
                            self.download(PMID,requests_pdf_url,savepath)
                            
                            
                        elif len(re.findall(r"/science/article/.*?[.]pdf", pdfpage)) > 0:
                            #通过sciencedirect网站
                            pdf_lianjie = re.findall(r"/science/article/.*?[.]pdf", pdfpage)
                            
                            requests_pdf_url = "https://www.sciencedirect.com"+pdf_lianjie[0]
                            
                            self.record = self.pubmed_info(PMID)
                            self.download(PMID,requests_pdf_url,savepath)                            
                            
                        elif len(tree_pdf.xpath('//*[@id="downloadPDFURL"]/@href')) > 0:
                            #通过https://www.spandidos-publications.com/网站下载文献
                            pdf_lianjie = tree_pdf.xpath('//*[@id="downloadPDFURL"]/@href')
                            
                            requests_pdf_url = "https://www.spandidos-publications.com/"+pdf_lianjie[0]
                            
                            self.record = self.pubmed_info(PMID)
                            self.download(PMID,requests_pdf_url,savepath)
                        
                        else:
                            pdf_lianjie = re.findall("https:.*?.pdf", pdfpage)
                            if len(pdf_lianjie) > 0:
                                requests_pdf_url = pdf_lianjie[0]
                                self.record = self.pubmed_info(PMID)
                                self.download(PMID,requests_pdf_url,savepath)
        
        self.success = []
        for filename in os.listdir(savepath):
            filename = filename.split(".")[0]
            self.success.append(filename)
        
        self.real_success = list(set(self.unique_data_y).intersection(set(self.success))) 
        
        self.no_data = list(set(self.unique_data_y).difference(set(self.success)))
            
        for PMID in self.no_data:
            self.record["PMID"].append(PMID)
            self.record["title"].append("空")
            self.record["journal"].append("空")
            self.record["doi"].append("空")
            self.record['是否成功下载'].append("否")
        
        
        
        #record = {"aa":[1,2,3],"bb":[4,5,6]}
            
        self.result_data_Text.insert(END,"不重复PMID："+ str(len(self.unique_data_y)) +"\n" +
                                         "目标文件夹下成功获得：" + str(len(self.real_success)) + "\n" +
                                         "未成功下载：" + str(len(self.no_data))+ "\n" +
                                         "详细情况请查看工作目录下的result.txt文件\n")
        self.init_window_name.update()        
        
        file = open('result.txt', 'w',encoding = "utf-8") 
        #将字典写入
        for i in self.record.keys():
            file.write(str(i)+"\t")
        file.write("\n")
        PMID_sum = len(list(self.record.values())[0])
        
        if PMID_sum > 0:
            for i in range(0,PMID_sum):
                for v in list(self.record.values()):
                    file.write(str(v[i])+"\t")
                file.write("\n")
                # 注意关闭文件
        file.close()

        
        """result = DataFrame(self.record)
        result.to_csv('result.csv',index = False)"""
        """ print("不重复PMID："+ str(len(unique_data_y)) +"\n" +
              "目标文件夹下成功获得：" + str(len(real_success)) + "\n" +
              "未成功下载：" + str(len(no_data))+ "\n" +
              "详细情况请查看工作目录下的result.csv文件\n")  """
        

        


def gui_start():
    init_window = Tk()              #实例化出一个父窗口
    ZMJ_PORTAL = DOC_GUI(init_window)
    # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()
    init_window.mainloop()          #父窗口进入事件循环，保持窗口运行，否则界面不展示


gui_start()






