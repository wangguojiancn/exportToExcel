# exportToExcel

###   导出数据到EXCEL(包含图片)

exportToExcel(data1,data2,Obj)

data1: 数组 表格头数据 如 ['a','b','c']

data2: 数组 表格主内容数据
	- 格式一:数组项为子数组 且与 data1内容一一对应  如 [['a1','a2','a3'],['b1','b2','b3']]
	- 格式二:数组项为子对象 对象中的键值对中的值和data1 一一对应  如 [{'a1':1,'a2':2,'a3':3},{'b1':1,'b2':2,'b3':3}]

Obj: 对象 可设置对应参数
	- filename : 文件名
	- sheetName : sheet名
	- width : 图片单元格宽度
	- height : 单元格高度