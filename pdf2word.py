import os
import sys
from pdf2docx.main import parse

# ドラッグ＆ドロップでパスを取得する
pdf_path = sys.argv[1]
# 拡張子の変換する
docx_path = pdf_path.split(".")[0] + ".docx"

parse(pdf_path, docx_path,start=1, end=2)

# def createcsv_results(pdf_faaaaa) :
#     txtlist = os.listdir('./data/results/txt')

#     for txtname in txtlist:	
#         docx_path = pdf_path.split(".")[0] + ".docx"
#         if not os.path.exists(docx_path):
#             txtfile = open("./data/results/txt/" + txtname ,"r")
#             csvfile = open(docx_path,"w")
#             self.parse_result(txtfile,csvfile)
                
#                 csvfile.close()
#                 txtfile.close()
