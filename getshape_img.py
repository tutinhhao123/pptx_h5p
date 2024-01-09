
#import js2py
import xml.etree.ElementTree as ET
import os
import xmltodict
import edit_colon2_ as edit
import json
import xml.dom.minidom
#import xml2json as xml2json
#import pyutil

# Array_obj = data['p:sld']["p:cSld"]["p:spTree"]["p:sp"]["p:txBody"]["a:p"]['a:p']['a:r']['a:t']
# Array_obj = data['p:sld']["p:cSld"]["p:spTree"]["p:sp"]

class Create_obj_p():          # leave this empty
    def __init__(self,text,x,y,cx,cy,size):   # constructor function using self
        self.text = text  # variable using self.
        #self.color = color  # variable using self
        self.x = x
        self.y = y
        self.cx = cx
        self.cy = cy
        self.size=size
def get_text_ngoai(root):
    j=0
    list = []
    text=""

    for i in root.findall('./p_cSld/p_spTree/p_grpSp'):
        # print(i.tag,i.attrib,i.text)

        #print(i.tag, i.attrib, i.text)
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t'):
            #print(x.tag, x.attrib, x.text)
            j += 1
            if(x.text!="" and x.text!=""):

                text = text + x.text
            print(j)
        #list.append(Create_obj_p(text, 3))

        print("textlay duoc",text)
        a = Create_obj_p('', '', '', '', '', '')

        if(text!=" " and text!=""):
            a.text=text
            list.append(a)
            print('a.text=',a.text)
        if(a.text=="" or a.text==" "):
            print('atexxt rong')

        j = 0
        text=""
    for i in list:
        print('value list0', i.text)

    return list
def get_text_trong(root):
    j = 0
    list = []
    text = ""
    for i in root.findall('./p_cSld/p_spTree'):
        # print(i.tag,i.attrib,i.text)

        print(i.tag, i.attrib, i.text)
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t'):
            # print(x.tag, x.attrib, x.text)
            j += 1
            if (x.text != "" and x.text != ""):
                text = text + x.text
            print(j)
            # list.append(Create_obj_p(text, 3))

        print("textlay duoc", text)
        a = Create_obj_p('', '', '', '', '', '')

        if (text != " " and text != ""):
            a.text = text
            list.append(a)
            print('a.text=', a.text)
        if (a.text == "" or a.text == " "):
            print('atexxt rong')

        j = 0
        text = ""
    for i in list:
        print('value list0', i.text)

    return list
def long_text(listtextngoai,listtexttrong):
    #listtrong - > listngoai
    for i in listtextngoai:
        print("print text trong:",i.text)
        listtexttrong.append(i)
    for i in listtexttrong:
        print("gia tri text listngoai:",i.text)
    return listtexttrong
def long_text_X_Y(listtextngoai,listtexttrong):
    #listtrong - > listngoai
    for i in listtextngoai:
        print("print text trong:",i.x)
        listtexttrong.append(i)
    for i in listtexttrong:
        print("gia tri text listngoai:",i.x)
    return listtexttrong
def finallyoftext(list):
    for i in list:
        print("value i",i.text)
    return list
def get_x_y__cxcy_text_trong(root):
    list=[]

    for i in root.findall('./p_cSld/p_spTree'):
        # print(i.tag,i.attrib,i.text)

        print(i.tag, i.attrib, i.text)
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../.../.../...'):
            # print(x.tag, x.attrib, x.text)
            for zz in x.findall('.p_spPr/a_xfrm/'):

                print('value element ', zz.text, zz.tag, zz.attrib)
            a = Create_obj_p('', '', '', '', '', '')
            for zz in x.findall('.p_spPr/a_xfrm/a_off'):
                #, zz.text, zz.tag
                print("value zz", zz.text, zz.tag, zz.attrib)
                a.y = zz.attrib['y']
                a.x = zz.attrib['x']
            for zz in x.findall('.p_spPr/a_xfrm/a_ext'):
                a.cx = zz.attrib['cx']
                a.cy = zz.attrib['cy']


            list.append(a)
    checkRT_list(list)
    return list
def get_x_y__cxcy_text_ngoai(root):
    list=[]

    for i in root.findall('./p_cSld/p_spTree/p_grpSp'):
        # print(i.tag,i.attrib,i.text)

        print(i.tag, i.attrib, i.text)
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../.../.../...'):
            # print(x.tag, x.attrib, x.text)
            for zz in x.findall('.p_spPr/a_xfrm/'):

                print('value element ', zz.text, zz.tag, zz.attrib)
            a = Create_obj_p('', '', '', '', '', '')
            for zz in x.findall('.p_spPr/a_xfrm/a_off'):
                #, zz.text, zz.tag
                print("value zz", zz.text, zz.tag, zz.attrib)
                a.y = zz.attrib['y']
                a.x = zz.attrib['x']
            for zz in x.findall('.p_spPr/a_xfrm/a_ext'):
                a.cx = zz.attrib['cx']
                a.cy = zz.attrib['cy']


            list.append(a)
    checkRT_list(list)
    return list
def checkRT_list(aa):
    for i in aa:
        print(i.x, i.y , i.cx,i.cy )

def fullobj_text(listtext,list_xy):
    list=[]

    # for i in listtext:
    #     a = Create_obj_p('', '', '', '', '', '')
    #
    #     a.text=i.text
    for f, b in zip(listtext, list_xy):
        #print(f.text, b.x)
        f.x=b.x
        f.y=b.y
        f.cx=b.cx
        f.cy=b.cy
    return listtext
def checkfull_text_and_xy(list):
    for i in list:
        print("x-y-cx-cy",i.x,i.y,i.cx,i.cy,'text',i.text)
        print()
    return list
def lay_all_objtext_infilexml(folder,title):
    aa, bb, cc, dd = edit.edit_name_xml2txt(title, folder)
    edit.edittext_(aa)
    xx, yy, zz, gg=edit.edit_name_txt2xml(title, folder)
    # xml2json.xml_2_json_getfromfile3()
    tree = ET.parse(xx)
    root = tree.getroot()
    print(tree)
    textngoai = get_text_ngoai(root)
    texttrong = get_text_trong(root)
    a = long_text(textngoai, texttrong)

    finallytext = finallyoftext(a)
    # b=get_x_y__cxcy_text_trong(root)
    # fullobj_text(a,b)
    # checkRT_list(a)
    # get_text_trong(root)
    ngoai = get_x_y__cxcy_text_ngoai(root)
    trong = get_x_y__cxcy_text_trong(root)
    x_y = long_text_X_Y(ngoai, trong)
    fullobj_list = fullobj_text(listtext=finallytext, list_xy=x_y)
    xyz=checkfull_text_and_xy(fullobj_list)
    return xyz
if __name__ == "__main__":
    # a,b,c,d=edit.edit_name_xml2txt('slide5',r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json')
    # edit.edittext_(a)
    # edit.edit_name_txt2xml('slide5',r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json')
    # #xml2json.xml_2_json_getfromfile3()
    # tree= ET.parse(r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json\slide5.xml')
    # root=tree.getroot()
    # print(tree)
    # textngoai=get_text_ngoai(root)
    # texttrong=get_text_trong(root)
    # a=long_text(textngoai,texttrong)
    #
    # finallytext=finallyoftext(a)
    # #b=get_x_y__cxcy_text_trong(root)
    # #fullobj_text(a,b)
    # #checkRT_list(a)
    # #get_text_trong(root)
    # ngoai=get_x_y__cxcy_text_ngoai(root)
    # trong=get_x_y__cxcy_text_trong(root)
    # x_y=long_text_X_Y(ngoai,trong)
    # fullobj_list=fullobj_text(listtext=finallytext,list_xy=x_y)
    # checkfull_text_and_xy(fullobj_list)
    #fullobj_text()
    #/p_txBody/a_p/a_r/a_t
    #p_grpSp/p_sp/a_p/a_r/a_t

    #size= p_spPr/a_xfrm -> a_off or a_ext


    #ET.
    a=lay_all_objtext_infilexml(folder=r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep\a1',title='slide5')
    for i in a:
        print("gia tri cuoi cung",i.text,i.x,i.y,i.cx,i.cy)