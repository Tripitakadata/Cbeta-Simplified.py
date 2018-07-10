#
# Cbeta 标点数据统计,注意本程序仅用于
#
import os  
import xlrd
import xlsxwriter
import zhon.hanzi as ZhongHanzi
import zhon.pinyin as ZhongPinyin
import zhon.zhuyin as ZhongZhuyin
import zhon.cedict as ZhongCedict

punctuation_set='。.？?!＂＃＄％＆＇（）＊＋，－／：；＜＝＞＠［＼］＾＿｀｛｜｝～｟｠｢｣､、〃》「」『』【】〔〕〖〗〘〙〚〛〜〝〞〟〰〾〿–—‘’‛“”„‟…‧﹏'
string_set='abcdefghijklmnopkrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
number_set='0123456789'
file_dir="/home/z/tripitaka_data/006.自有数据"

#tripitakalist=["经文代号","经文名称","字符总数","汉字总数","汉字比率","标点总数","句号比率","问号比率","感叹号比率","冒号比率","引号比率","逗号比率","顿号比率","汉字标点数"]
tripitakalist=["经文代号","字符总数","汉字总数","汉字比率","标点总数","句号比率","问号比率","感叹号比率","冒号比率","引号比率","逗号比率","顿号比率","汉字标点数"]

tripitakaTpye=[]
Total_punctuation=0
Total_character=0
workbook = xlsxwriter.Workbook('藏经-自有数据-标点统计.xls')
worksheet = workbook.add_worksheet()


row = 0
col = 0

for temp in tripitakalist:
    worksheet.write(row, col, temp)
    col += 1

col = 0
row = 1
for root, dirs, files in os.walk(file_dir):

    for file_temp in sorted(files):
        Total_punctuation=0
        Total_character=0  
        if file_temp == "Readme.txt":
            pass
        else:
            #写入Cbeta所有经文的名字
            worksheet.write(row, col, file_temp.split('.')[0]) 
            #print(root)           
            FileData = open(root+'/'+file_temp, 'r', encoding='utf-8')
            lines = FileData.readlines()#读取全部内容
            Juan=''.join(lines)
            #Juan=''.join(lines).split('<juan>')[1]          
            #Title=''.join(Juan).split('<p>')[0]
            #worksheet.write(row, col+1,Title) 
            #################################################
            ####总标点符号统计
            for i in punctuation_set:
                Total_punctuation+=Juan.count(i)
            Total_character=Total_punctuation
            for i in string_set:
                Total_character+=Juan.count(i)
            for i in number_set:
                Total_character+=Juan.count(i)
            for i in Juan:
                if i >= u'\u4e00' and i<=u'\u9fa5':
                    Total_character+=1  
            worksheet.write(row, 1, Total_character+Total_punctuation)
            worksheet.write(row, 2, Total_character)
            worksheet.write(row, 3, Total_character/(Total_character+Total_punctuation))          
            worksheet.write(row, 4, Total_punctuation) 
            Total_punctuation=max(Total_punctuation,1)
            worksheet.write(row, 5, (Juan.count('。')+Juan.count('.'))/Total_punctuation) 
            worksheet.write(row, 6, (Juan.count('？')+Juan.count('?'))/Total_punctuation)
            worksheet.write(row, 7, Juan.count('！')/Total_punctuation) 
            worksheet.write(row, 8, Juan.count('：')/Total_punctuation)
            worksheet.write(row, 9, (Juan.count('「')+Juan.count('」')+Juan.count('『')+Juan.count('』')+Juan.count('”')+Juan.count('“'))/Total_punctuation)
            worksheet.write(row, 10, Juan.count('，')/Total_punctuation)
            worksheet.write(row, 11, Juan.count('、')/Total_punctuation)
            worksheet.write(row, 12, Total_character/max(1,(Juan.count('。')+Juan.count('.')+Juan.count('？')+Juan.count('?')+Juan.count('！'))))


            FileData.close()
            row += 1
#关闭文件
workbook.close()


