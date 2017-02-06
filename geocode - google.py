#-*- coding:utf-8 -*-
import xlrd
from xlutils.copy import copy
from xlwt import *
from geopy.exc import GeocoderTimedOut, GeocoderServiceError
from geopy.geocoders import *
import os

#格式要求：数据从第四行开始，数据结束后不能有多余行
filename = raw_input('输入文档名：')
data = xlrd.open_workbook(filename, formatting_info=True)
table = data.sheets()[0]
#取地址信息
addresscol = []
#Primary-and-Special-Schools-List
addresscol1 = table.col_values(4)
addresscol2 = table.col_values(5)
addresscol3 = table.col_values(6)
addresscol4 = table.col_values(7)
#stats_FULL_SCHOOL_LIST
# addresscol1 = table.col_values(3)
# addresscol2 = table.col_values(4)
for each in range(3,len(addresscol1)):
	print addresscol
	try:
		addresscol.append(addresscol1[each]+' '+addresscol2[each]+' '+addresscol3[each]+' '+addresscol4[each])
		#addresscol.append(addresscol1[each] + ' ' + addresscol2[each])
	except TypeError:
		addresscol.append(str(addresscol1[each])+' '+str(addresscol2[each])+' '+str(addresscol3[each])+' '+str(addresscol4[each]))
		#addresscol.append(str(addresscol1[each]) + ' ' + str(addresscol2[each]))

print len(addresscol)
#用于存放坐标
coordinate = []
#google api key
apikey = [	  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************',
		  '***************************************']
#num用于记载正在使用第几个api key
num = 0
#po用于记载内置循环到乐地址列表的哪一项，实现更换key后从断点处继续工作
po = 0

#通过函数递归调用的形式实现key自动更换
def geo_code(apikey, coordinate, addresscol, po, num):
	geolocator = GoogleV3(api_key=apikey[num])
	#转码并进行地理编码
	for i in range(po,len(addresscol)):
		addresscol[i].encode('utf-8')
		try:
			if len(coordinate) < len(addresscol):
				#限定查询国家为爱尔兰，控制超时上限缓解网络缓慢问题
				location = geolocator.geocode(addresscol[i], timeout=30, components={'country':'Ireland'},exactly_one=True)
				coordinate.append((location.latitude, location.longitude))
				print (location.latitude, location.longitude)
			else:
				print "编码完成！"
				break
		except AttributeError:
			print '地理编码失败！'
			coordinate.append((0,0))
		except GeocoderTimedOut:
			po = i
			print "连接超时，检查网络链接或更改timeout参数"
			geo_code(apikey, coordinate, addresscol, po, num)
		except GeocoderServiceError:
			#用于key超过配额处理
			po = i
			if num < len(apikey)-1:
				num += 1
				print "key changed"
				geo_code(apikey, coordinate, addresscol, po, num)
			else:
				print "apikey不足，停止编码进行输出"
				break


	return coordinate

geo_code(apikey, coordinate, addresscol, po, num)

#尝试创建结果文件夹，若文件夹已存在，直接更换工作目录至结果文件夹
try:
	os.mkdir('result')
	os.chdir('result')
except:
	os.chdir('result')
#拷贝原有excel数据
newdata = copy(data)
sheet = newdata.get_sheet(0)
#在拷贝数据上进行写入数据操作
for i in range(3,table.nrows):
	# Primary-and-Special-Schools-List
	sheet.write(i,17,coordinate[i-3][0])
	sheet.write(i,18,coordinate[i-3][1])
	# stats_FULL_SCHOOL_LIST
	# sheet.write(i, 9, coordinate[i - 3][0])
	# sheet.write(i, 10, coordinate[i - 3][1])
#保存结果表格
newdata.save(filename)

