'''
Date: 3 September
Module: myxls.py
Author: Prateek Mehta
Contact: prateek.mehta1992@gmail.com
'''


import xlrd
import webbrowser
import re
import html

##############################################################################################################
folder = r"D:\Python33\python_2013\xls"
_infile = "in.xls"
_outfile = "out.html"
inlocation = "{}\\{}".format(folder,_infile)
outlocation = "{}\\{}".format(folder,_outfile)
frh = xlrd.open_workbook(inlocation)
sheets = frh.sheet_names()
sheetdata = {}
location_indices = []
##############################################################################################################
datepattern = re.compile(r"date:\s([0-9]+)/([0-9]+)/.*",re.I)
def write(line,fh): print(line,end="",file=fh)
##############################################################################################################
for name in sheets:
	sheet = frh.sheet_by_name(name)
	rawdatetime = sheet.cell_value(2,1)
	match = re.search(datepattern,rawdatetime)
	sheetdata[str(match.group(1)).lstrip("0")] = sheet
##############################################################################################################
# http://www.python.org/dev/peps/pep-0008/#descriptive-naming-styles
# class vs. class_
for e,key_ in enumerate(sheetdata.keys()):
##############################################################################################################
	sheet = sheetdata[key_]
	i = 0
	locations = []
##############################################################################################################
	while(i < sheet.nrows):
		(p,n) = (sheet.cell_value(i,0),sheet.cell_value(i,1))
		i += 1
		prevpattern = re.compile(r"^[0-9]+")
		next_pattern_a = re.compile(r"^([a-z]+[\s]?[0-9]?).*$",re.I)
		next_pattern_b = re.compile(r"^([a-z]+[\.]?[\s]?-[\s]?[a-z]+[\s]?[0-9]?).*$",re.I)
		next_pattern_c = re.compile(r"^([a-z]+[\.]?[\s]+[a-z]+[\s]?[a-z]+[\s]?[a-z]+).*$",re.I)
		if re.search(prevpattern,str(p)):
			if e==0: location_indices.append(i-1)
			match_a = re.search(next_pattern_a,str(n))
			match_b = re.search(next_pattern_b,str(n))
			match_c = re.search(next_pattern_c,str(n))
			if(match_a and match_b):
				locations.append(str(match_b.group(1)).strip() + "~")
				continue
			if(match_a and match_c):
				locations.append(str(match_c.group(1)).strip() + "~")
				continue
			if(match_a):
				locations.append(str(match_a.group(1)).strip() + "~")
	sheetdata[key_] = {}
	sheetdata[key_]["locations"] = []
	sheetdata[key_]["locations"].extend(locations)
# for key_ in sheetdata.keys(): print("{}: ".format(key_).ljust(5), sheetdata[key_]["locations"])
# print("Indices: {}".format(location_indices))
##############################################################################################################
# 3,4,5 VSAT-CANOPY
# 6,7,8 MPLS-RELIANCE
# 9,10,11 MPLS-BSNL
# Status, Availability, RoundTripDelay
'''
sheetdata[key_]:{ # key_ = date
	locations:[a,b,c,] # for location in date["locations"]: pass
	data:{
		VSAT-CANOPY:[(a1,b1,c1),(a2,b2,c2),(a3,b3,c3),] # for s,a,r in date["data"]["VSAT-CANOPY"]: pass
		MPLS-RELIANCE:[(a1,b1,c1),(a2,b2,c2),(a3,b3,c3),] # for s,a,r in date["data"]["MPLS-RELIANCE"]: pass
		MPLS-BSNL:[(a1,b1,c1),(a2,b2,c2),(a3,b3,c3),] # for s,a,r in date["data"]["MPLS-BSNL"]: pass
	}
}
'''
data_a_name = "VSAT-CANOPY"
data_b_name = "MPLS-RELIANCE"
data_c_name = "MPLS-BSNL"
data_a = data_b = data_c = [] # VSAT-CANOPY, MPLS-RELIANCE, MPLS-BSNL
for name in sheets:
	sheet = frh.sheet_by_name(name)
	for i in location_indices:
		a = (sheet.cell_value(i,2),sheet.cell_value(i,3),sheet.cell_value(i,4))
		b = (sheet.cell_value(i,5),sheet.cell_value(i,6),sheet.cell_value(i,7))
		c = (sheet.cell_value(i,8),sheet.cell_value(i,9),sheet.cell_value(i,10))
		data_a.append(a)
		data_b.append(b)
		data_c.append(c)
for key_ in sheetdata.keys():
	sheetdata[key_]["data"] = {}
	sheetdata[key_]["data"][data_a_name] = data_a
	sheetdata[key_]["data"][data_b_name] = data_b
	sheetdata[key_]["data"][data_c_name] = data_c
for key_ in sheetdata.keys():
	print(key_,end=">>>\n")
	for key__ in sheetdata[key_].keys():
		print(key__,sheetdata[key_][key__])
##############################################################################################################
fwh = open(outlocation,"w+t")
head = \
'''
<head>
<style>
html,body,div {
	padding: 0;
	margin: 0;
}
body {
	background: -moz-linear-gradient(top, rgba(255,255,255,1) 0%, rgba(255,255,255,0) 100%);
	background: -webkit-linear-gradient(top, rgba(255,255,255,1) 0%,rgba(255,255,255,0) 100%);
	background: -o-linear-gradient(top, rgba(255,255,255,1) 0%, rgba(255,255,255,0) 100%);
	background: -ms-linear-gradient(top, rgba(255,255,255,1) 0%, rgba(255,255,255,0) 100%);
	background: linear-gradient(to bottom, rgba(255,255,255,1) 0%, rgba(255,255,255,0) 100%);
	background-image: url("bg.jpg");
	background-repeat: no-repeat;
	background-position: center center;
}
div.main {
	position: absolute;
	display: block;
	box-shadow: 0px 0px 20px rgba(200, 200, 200, 0.3);
	background: rgba(200, 200, 200, 0.4);
	border: 2px dashed white;
	width: 450px;
	height: 550px;
	border-radius: 4px;
	margin: auto auto;
	padding: 10px;
	top: 0;
	bottom: 0;
	left: 0;
	right: 0;
	transition: all 500ms ease 500ms;
}
div.main:hover {
	box-shadow: 0px 0px 20px rgba(150, 150, 150, 0.4);
	background: rgba(180, 180, 180, 0.5);
	-webkit-transition: background-color 500ms ease;
	-moz-transition: background-color 500ms ease;
	-o-transition: background-color 500ms ease;
	-ms-transition: background-color 500ms ease;
	transition: background-color 500ms ease;
}
div.main:hover div.row {
	text-shadow: 0px 0px 30px white;
	-webkit-transition: text-shadow 500ms ease;
	-moz-transition: text-shadow 500ms ease;
	-o-transition: text-shadow 500ms ease;
	-ms-transition: text-shadow 500ms ease;
	transition: text-shadow 500ms ease;
}
div.row {
	font-family: Arial, Helvetica, Sans-Serif;
	text-shadow: 0px 0px 10px white;
	color: white;
	font-size: 17px;
	font-weight: bold;
	border-bottom: 2px dotted black;
	transition: all 500ms ease 500ms;
}
div.row::before {
	content: "+ ";
}
</style>
<link rel='shortcut icon' href='http://static.pixdip.com/favicon.ico' type='image/x-icon'>
<title>myxls</title>
</head>
'''
write(head,fwh)
write("<body>",fwh)
write("<div class='main'>",fwh)
keys = list(sheetdata.keys())
set_ = sheetdata[keys[0]]["locations"]
'''
tkinter code
'''
# print("Duplicates: ", list(set([x for x in set_ if set_.count(x) > 1])))


write("<select>",fwh)
for i,location in enumerate(sheetdata[keys[0]]["locations"]):
	write("<option value='{0}'>{1}</option>".format(i,location),fwh)
	# write("<div class='row'></div>",fwh)
write("</select>",fwh)


write("</div>",fwh)
write("</body>",fwh)
fwh.close()
##############################################################################################################
# webbrowser.open_new(outlocation)