import copy
#import js2py
import xml.etree.ElementTree as ET
import os
import traceback
import xmltodict
import edit_colon2_ as edit
import json

import xml.dom.minidom
#import xml2json as xml2json
#import pyutil
import nhap5_usedict_addobj
# Array_obj = data['p:sld']["p:cSld"]["p:spTree"]["p:sp"]["p:txBody"]["a:p"]['a:p']['a:r']['a:t']
# Array_obj = data['p:sld']["p:cSld"]["p:spTree"]["p:sp"]

class Create_obj_p():          # leave this empty
    def __init__(self,text,x,y,cx,cy,size,color):   # constructor function using self
        self.text = text  # variable using self.
        #self.color = color  # variable using self
        self.x = x
        self.y = y
        self.cx = cx
        self.cy = cy
        self.size=size
        self.color = color
class Create_obj_theme():          # leave this empty
    def __init__(self,tag_color,value_color):   # constructor function using self
        self.tag_color = tag_color  # variable using self.
        #self.color = color  # variable using self
        self.value_color = value_color
class Create_obj_xmlrel():          # leave this empty
    def __init__(self,Id,Target):   # constructor function using self
        self.Target = Target  # variable using self.
        #self.color = color  # variable using self
        self.Id = Id
class Create_obj_shapeandimg():
    def __init__(self,rId,x,y,cx,cy,prst,img,color ):
        self.rId = rId  # variable using self.
        # self.color = color  # variable using self
        self.x = x
        self.y = y
        self.cx = cx
        self.cy = cy
        self.prst = prst
        self.img = img
        self.color = color


def get_text_ngoai(root,root_oftheme):
    j=0
    list = []
    text=""
    #COLOR
    color=get_colorr(root_oftheme)
    for i in color:
        print('i tag_value in textngoai',i.value_color,'i tag_color',i.tag_color)
    # END COLOR
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
        a = Create_obj_p('', '', '', '', '', '','')
        # size
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../a_rPr'):
            print('value element ', x.text, x.tag,x.attrib)#, x.attrib['sz'], x.attrib)
            #--------- tạm thời không lấy size chữ. x.attrib['sz'] 9 27 2020
            #a.size=x.attrib['sz']
        #------------



        # color
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/...//a_solidFill/'):
            print('value element ', x.text, x.tag, x.attrib['val'])
            print(1)
            for i in color:
                if str(x.attrib['val'])==str(i.tag_color):
                    a.color=i.value_color
                else:
                    if not x.attrib['val']:
                        a.color='000001'
                    else:

                        a.color=x.attrib['val']
                    #a.color=x.attrib['val']
        if(text!=" " and text!=""):
            a.text=text
            list.append(a)
            print('a.text=',a.text)
        if(a.text=="" or a.text==" "):
            print('atexxt rong')

        j = 0
        text=""
    # for i in list:
    #     print('value list0 color', i.text, 'size', i.size, 'color', i.color)
    return list
def get_text_trong(root,root_oftheme):
    j = 0
    list = []
    text = ""
    listtheme=[]
    # COLOR
    color = get_colorr(root_oftheme)
    for i in color:
        print('i tag_value in texttrong',i.value_color,'i tag_color',i.tag_color)
    # END COLOR
    for i in root_oftheme.findall('./p_cSld/p_spTree'):
        # print(i.tag,i.attrib,i.text)

        print(i.tag, i.attrib, i.text)
        #listtheme.append(Create_obj_theme())
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
        a = Create_obj_p('', '', '', '', '', '','')
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../a_rPr'):
            print('value element ', x.text, x.tag)#, x.attrib['sz'],x.attrib)
            # --------- tạm thời không lấy size chữ. x.attrib['sz'] 9 27 2020
            # a.size=x.attrib['sz']

        #color

        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/...//a_solidFill/'):
            print('value element color ', x.text, x.tag,x.attrib['val'])
            for ii in color:
                if str(x.attrib['val'])==str(ii.tag_color):
                    a.color=ii.value_color
                else:
                    if not x.attrib['val']:
                        a.color='000001'
                    else:

                        a.color=x.attrib['val']

        if (text == " " or text != ""):
            a.text = text
            list.append(a)
            print('a.text=', a.text)
        if (a.text == "" or a.text == " "):
            print('atexxt rong')

        j = 0
        text = ""
    # for i in list:
    #     print('value list0 color', i.text,'size',i.size,'color',i.color)

    return list
def long_text(listtextngoai,listtexttrong,root_oftheme):
    #listtrong - > listngoai
    print('#_------------start longtext-------------')
    color = get_colorr(root_oftheme)
    for i in listtextngoai:

        print("print text trong thuoc longtext va text ngoai:",i.text,i.color,i.size)
        if not i.color:
            i.color='000001'
        for j in color:
            print('gia tri i color[2:]', j.tag_color, j.tag_color[2:])
            if str(i.color)==str(j.tag_color[2:]):
                i.color=j.value_color
        listtexttrong.append(i)
    for i in listtexttrong:
        print("gia tri text listngoai longtext va texttrong:",i.text,i.color,i.size)
        if not i.color:
            i.color = '000001'
        for j in color:
            print('gia tri i color[2:]',j.tag_color,j.tag_color[2:])
            if str(i.color) == str(j.tag_color[2:]):
                i.color = j.value_color
    print('#_------------end longtext-------------')
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
        print("value i",i.text,i.size,i.color)
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
            a = Create_obj_p('', '', '', '', '', '','')
            for zz in x.findall('.p_spPr/a_xfrm/a_off'):
                #, zz.text, zz.tag
                print("value zz", zz.text, zz.tag, zz.attrib)
                a.y = zz.attrib['y']
                a.x = zz.attrib['x']
            for zz in x.findall('.p_spPr/a_xfrm/a_ext'):
                a.cx = zz.attrib['cx']
                a.cy = zz.attrib['cy']
            #for zz in x.findall('.p_spPr/a_xfrm/a_off'):

            list.append(a)
    checkRT_list(list)
    return list
def convertpixel_ems(pixel):
    #12 point -> 0,995846 ems
    #3000 pixel -> 2249.9982992137 point
    #1 inch = 914400 EMU, 1 mm = 36000 EMU.
    #1pt -> 0.35277777777778 mm
    # emu = 12 *1 *0.35277777777778*36000
    ems=''
    return ems
def get_x_y__cxcy_text_ngoai(root):
    list=[]

    for i in root.findall('./p_cSld/p_spTree/p_grpSp'):
        # print(i.tag,i.attrib,i.text)

        print(i.tag, i.attrib, i.text)
        for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../.../.../...'):
            # print(x.tag, x.attrib, x.text)
            for zz in x.findall('.p_spPr/a_xfrm/'):

                print('value element ', zz.text, zz.tag, zz.attrib)
            a = Create_obj_p('', '', '', '', '', '','')
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
        if (f.x[0] == '-'):
            f.x = str(0)
        if (f.y[0] == '-'):
            f.y = str(0)
        if (f.cx[0] == '-'):
            f.cx = str(0)
        if (f.cy[0] == '-'):
            f.cy = str(0)
        f.x = str((100.0 * float(f.x)) / 9144000.0)

        f.y = str((100.0 * float(f.y)) / 6858000.0)
        f.cx = str((100.0 * float(f.cx)) / 9144000.0)
        f.cy = str((100.0 * float(f.cy)) / 6858000.0)
    return listtext
def checkfull_text_and_xy(list):
    for i in list:
        print("finally checkx-y-cx-cy",i.x,i.y,i.cx,i.cy,'text',i.text,'size',i.size,'color',i.color)
        print()
    return list

def get_shape_obj_inrels_001(root):
    list = []


    # tree = ET.parse(
    #     r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\a1_copy\ppt\slides\_rels\slide5.xml.rels')
    # root = tree.getroot()
    #print(root)
    for i in root.findall('.//'):
        #zxxz = Create_obj_xmlrel('', '')
        # print(i.tag,i.attrib,i.text)
        print(i.attrib['Target'], i.attrib)
        list.append(Create_obj_xmlrel(Id=i.attrib['Id'],Target=i.attrib['Target']))

    checkRT_img_XMLRELS_list(list)

    return list
def checkRT_list(aa):
    for i in aa:
        print('check rt_list',i.x,i.y )
def checkRT_img_XMLRELS_list(aa):
    for i in aa:
        print('check target-id',i.Target[2:],i.Id )
        print('images'+i.Target[2:])
def get_shape_obj_img_shape_001_old(root):
    list_shape = []
    count=0
    for i in root.findall('./p_cSld/p_spTree'):
        # print(i.tag,i.attrib,i.text)

        # print(i.tag, i.attrib, i.text)
        for x in i.findall('./'):
            if x.tag == 'p_sp':
                # get x y
                aaa = Create_obj_shapeandimg('', '', '', '', '', '', '', '')
                for j in i.findall('./p_sp/p_spPr/a_xfrm/a_off'):
                    print(j.tag, j.attrib, j.attrib['x'], j.attrib['y'], j.text)

                    aaa.x = j.attrib['x']
                    aaa.y = j.attrib['y']
                    if (int(aaa.x) < 0):
                        aaa.x = 0
                    if (int(aaa.y) < 0):
                        aaa.y = 0
                # get cx cy
                for j in i.findall('./p_sp/p_spPr/a_xfrm/a_ext'):
                    print(j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)

                    aaa.cx = j.attrib['cx']
                    aaa.cy = j.attrib['cy']
                    if (int(aaa.cx) < 0):
                        aaa.cx = 0
                    if (int(aaa.cy) < 0):
                        aaa.cy = 0
                # get prstGeom - shape - rec -line . . .
                for j in i.findall('./p_sp/p_spPr/a_prstGeom'):
                    print(j.tag, j.attrib, j.attrib['prst'], j.text)  # j.attrib['cx'], j.attrib['cy'],
                    aaa.prst = j.attrib['prst']
                # get solid_Fill
                for j in i.findall('./p_sp/p_spPr/a_solidFill/a_srgbClr'):
                    print(j.tag, j.attrib, j.attrib['val'], j.text)  # j.attrib['val'] j.attrib['cx'], j.attrib['cy'],
                    aaa.color = j.attrib['val']
                list_shape.append(aaa)
                count=count+1
                print('gia tri index count:',count)

            if x.tag == 'p_pic':
                # get x y
                aaa = Create_obj_shapeandimg('', '', '', '', '', '', '', '')
                for j in i.findall('./p_pic/p_spPr/a_xfrm/a_off'):
                    print(j.tag, j.attrib, j.attrib['x'], j.attrib['y'], j.text)

                    aaa.x = j.attrib['x']
                    aaa.y = j.attrib['y']
                    if (int(aaa.x) < 0):
                        aaa.x = 0
                    if (int(aaa.y) < 0):
                        aaa.y = 0
                # get cx cy
                for j in i.findall('./p_pic/p_spPr/a_xfrm/a_ext'):
                    print(j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)

                    aaa.cx = j.attrib['cx']
                    aaa.cy = j.attrib['cy']
                    if (int(aaa.cx) < 0):
                        aaa.cx = 0
                    if (int(aaa.cy) < 0):
                        aaa.cy = 0
                # get prstGeom - shape - rec -line . . .
                for j in i.findall('./p_pic/p_spPr/a_prstGeom'):
                    print(j.tag, j.attrib, j.attrib['prst'], j.text)  # j.attrib['cx'], j.attrib['cy'],
                    aaa.prts = j.attrib['prst']
                # get Rid_img
                for j in i.findall('./p_pic/p_blipFill/a_blip'):  # /a_extLst
                    print(j.tag, j.attrib['r_embed'], j.attrib,
                          j.text)  # j.attrib['val'] j.attrib['cx'], j.attrib['cy'],
                    aaa.rId = j.attrib['r_embed']
                list_shape.append(aaa)
                count=count+1
                print('gia tri index count:',count)
    print("gia tri cua count :",count)
    return list_shape
def get_shape_obj_img_shape_001(root):
    list_shape = []
    count=0
    numb_a_xfrm=0
    for i in root.findall('./p_cSld/p_spTree'):
        # print(i.tag,i.attrib,i.text)

        # print(i.tag, i.attrib, i.text)
        for x in i.findall('./'):
            if x.tag == 'p_cxnSp':
                # get x y
                aaa = Create_obj_shapeandimg('', '', '', '', '', '', '', '')
                for j in x.findall('./p_spPr/a_xfrm/a_off'):
                    print('p_sp check get x y p_sp:',j.tag, j.attrib, j.attrib['x'], j.attrib['y'], j.text)
                    print('numb',numb_a_xfrm)
                    numb_a_xfrm=numb_a_xfrm+1
                    aaa.x = j.attrib['x']
                    aaa.y = j.attrib['y']
                    if (int(aaa.x) < 0):
                        aaa.x = 0
                    if (int(aaa.y) < 0):
                        aaa.y = 0

                    aaa.x=((100.0 * float(aaa.x)) / 9144000.0)
                    aaa.y = ((100.0 * float(aaa.y)) / 9144000.0)

                    # xx #= ((100.0 * (xx)) / 9144000.0)
                    # yy #= ((100.0 * (yy)) / 9144000.0)
                    # cxx #= ((100.0 * (cxx)) / 9144000.0)
                    # cyy #= ((100.0 * (cyy)) / 9144000.0)

                # get cx cy
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../a_ext'):
                    print(' p_sp a_ext check:',j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)

                    aaa.cx = j.attrib['cx']
                    aaa.cy = j.attrib['cy']
                    if (int(aaa.cx) < 0):
                        aaa.cx = 0
                    if (int(aaa.cy) < 0):
                        aaa.cy = 0
                aaa.cx = ((100.0 * float(aaa.cx)) / 6858000.0)
                aaa.cy = ((100.0 * float(aaa.cy)) / 6858000.0)
                if (int(aaa.cx) > 100):
                    aaa.cx = 100
                if (int(aaa.cy) > 100):
                    aaa.cy = 100
                # xx #= ((100.0 * (xx)) / 9144000.0)
                # yy #= ((100.0 * (yy)) / 9144000.0)
                # cxx #= ((100.0 * (cxx)) / 9144000.0)
                # cyy #= ((100.0 * (cyy)) / 9144000.0)
                # get prstGeom - shape - rec -line . . .
                #for j in x.findall('./p_spPr/a_prstGeom'):
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../.../a_prstGeom'):
                    print('p_sp check prst',j.tag, j.attrib, j.attrib['prst'], j.text)  # j.attrib['cx'], j.attrib['cy'],
                    aaa.prst = j.attrib['prst']
                # get solid_Fill
                #for j in x.findall('./p_spPr/a_solidFill/a_srgbClr'):
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../.../a_ln/a_solidFill/a_srgbClr'):

                    print('p_sp check get solid_Fill',j.tag, j.attrib, j.attrib['val'], j.text)  # j.attrib['val'] j.attrib['cx'], j.attrib['cy'],
                    aaa.color = j.attrib['val']
                list_shape.append(aaa)
                count=count+1
                print('gia tri index count p_sp:',count)
            if x.tag == 'p_sp':
                # get x y
                aaa = Create_obj_shapeandimg('', '', '', '', '', '', '', '')
                for j in x.findall('./p_spPr/a_xfrm/a_off'):
                    print('p_sp check get x y p_sp:',j.tag, j.attrib, j.attrib['x'], j.attrib['y'], j.text)
                    print('numb',numb_a_xfrm)
                    numb_a_xfrm=numb_a_xfrm+1
                    aaa.x = j.attrib['x']
                    aaa.y = j.attrib['y']
                    if (int(aaa.x) < 0):
                        aaa.x = 0
                    if (int(aaa.y) < 0):
                        aaa.y = 0
                    aaa.x=((100.0 * float(aaa.x)) / 9144000.0)
                    aaa.y = ((100.0 * float(aaa.y)) / 9144000.0)
                # get cx cy
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../a_ext'):
                    print(' p_sp a_ext check:',j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)

                    aaa.cx = j.attrib['cx']
                    aaa.cy = j.attrib['cy']
                    if (int(aaa.cx) < 0):
                        aaa.cx = 0
                    if (int(aaa.cy) < 0):
                        aaa.cy = 0
                    aaa.cx = ((100.0 * float(aaa.cx)) / 6858000.0)
                    aaa.cy = ((100.0 * float(aaa.cy)) / 6858000.0)
                    if (int(aaa.cx) > 100):
                        aaa.cx = 100
                    if (int(aaa.cy) >100):
                        aaa.cy = 100

                # get prstGeom - shape - rec -line . . .
                #for j in x.findall('./p_spPr/a_prstGeom'):
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../.../a_prstGeom'):
                    print('p_sp check prst',j.tag, j.attrib, j.attrib['prst'], j.text)  # j.attrib['cx'], j.attrib['cy'],
                    aaa.prst = j.attrib['prst']
                # get solid_Fill
                #for j in x.findall('./p_spPr/a_solidFill/a_srgbClr'):
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../.../a_solidFill/a_srgbClr'):

                    print('p_sp check get solid_Fill',j.tag, j.attrib, j.attrib['val'], j.text)  # j.attrib['val'] j.attrib['cx'], j.attrib['cy'],
                    aaa.color = j.attrib['val']
                list_shape.append(aaa)
                count=count+1
                print('gia tri index count p_sp:',count)
            numb_a_xfrm=0
            if x.tag == 'p_pic':
                # get x y
                aaa = Create_obj_shapeandimg('', '', '', '', '', '', '', '')
                for j in x.findall('./p_spPr/a_xfrm/a_off'):
                    print(j.tag, j.attrib, j.attrib['x'], j.attrib['y'], j.text)

                    aaa.x = j.attrib['x']
                    aaa.y = j.attrib['y']
                    if (int(aaa.x) < 0):
                        aaa.x = 0
                    if (int(aaa.y) < 0):
                        aaa.y = 0
                    aaa.x=((100.0 * float(aaa.x)) / 9144000.0)
                    aaa.y = ((100.0 * float(aaa.y)) / 9144000.0)
                # get cx cy
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../a_ext'):
                #for j in x.findall('./p_spPr/a_xfrm/a_ext'):
                    print('p pic check a_ext:',j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)

                    aaa.cx = j.attrib['cx']
                    aaa.cy = j.attrib['cy']
                    if (int(aaa.cx) < 0):
                        aaa.cx = 0
                    if (int(aaa.cy) < 0):
                        aaa.cy = 0
                    aaa.cx = ((100.0 * float(aaa.cx)) / 6858000.0)
                    aaa.cy = ((100.0 * float(aaa.cy)) / 6858000.0)
                if (int(aaa.cx) > 100):
                    aaa.cx = 0
                if (int(aaa.cy) > 100):
                    aaa.cy = 100
                # get prstGeom - shape - rec -line . . .
                #./p_spPr/a_xfrm/a_off/.../a_ext
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../.../a_prstGeom'):
                    print('p pic check prst:',j.tag, j.attrib, j.attrib['prst'], j.text)  # j.attrib['cx'], j.attrib['cy'],
                    aaa.prts = j.attrib['prst']
                # get Rid_img
                for j in x.findall('./p_spPr/a_xfrm/a_off/.../.../.../p_blipFill/a_blip'):#:.../p_blipFill/a_blip'):  # /a_extLst
                   #print('check rid:',j.tag,j.attrib)
                    print('p pic r_embed check ',j.tag, j.attrib['r_embed'], j.attrib,
                          j.text)  # j.attrib['val'] j.attrib['cx'], j.attrib['cy'],
                    aaa.rId = j.attrib['r_embed']
                list_shape.append(aaa)
                count=count+1
                print('gia tri index count:',count)
    print("gia tri cua count :",count)
    for z in list_shape:
        # print(type((z.x)), type(z.y), z.cx, z.cy)
        # xx = float(copy.copy(z.x))
        # yy = float(copy.copy(z.y))
        # cxx = float(copy.copy(z.cx))
        # cyy=float(copy.copy(z.cy))
        #xx #= ((100.0 * (xx)) / 9144000.0)
        #yy #= ((100.0 * (yy)) / 9144000.0)
        #cxx #= ((100.0 * (cxx)) / 9144000.0)
        #cyy #= ((100.0 * (cyy)) / 9144000.0)

        #z.x = xx#((100.0 * int(z.x)) / 9144000.0)

        #z.y = yy#((100.0 * int(z.y)) / 6858000.0)
        #z.cx = cxx#((100.0 * int(z.cx)) / 9144000.0)
        #z.cy = cyy#((100.0 * int(z.cy)) / 6858000.0)
       print(z.x,z.y,z.cx,z.cy)
        # z.x = ((100.0 * int(z.x)) / 9144000.0)
        #
        # z.y = ((100.0 * int(z.y)) / 6858000.0)
        # z.cx = ((100.0 * int(z.cx)) / 9144000.0)
        # z.cy = ((100.0 * int(z.cy)) / 6858000.0)
    return list_shape

def check_list_shape(list_shape):
    print('check list shape')
    for i,dataa in enumerate(list_shape):
        print(i)
        print('dataa.rId',dataa.rId,'\ndataa.x',dataa.x,'dataa.y',dataa.y,'\ndataa.cx',dataa.cx,'dataa.cy',dataa.cy,'\ndataa.prst',dataa.prst,'dataa.img',dataa.img,'\ndataa.color',dataa.color)
        if dataa.rId!='':
            a=0

        else:
            print('nonee')
            # self.prst = prst
            # self.img = img
            # self.color = color

            #print('dataa.x',dataa.x,'dataa.y',dataa.y,'dataa.cx',dataa.cx,'dataa.cy',dataa.cy,'dataa.x',dataa.x,'dataa.prst',dataa.prst,'dataa.img',dataa.img,'dataa.color',dataa.color)
    return list_shape

def lay_all_objtext_infilexml(folder,title):
    try:
        aa, bb, cc, dd ,link_rel= edit.edit_name_xml2txt(title, folder)
        edit.edittext_(aa)
        edit.edittext_(link_rel)

        xx, yy, zz, gg,link_rel_new=edit.edit_name_txt2xml(title, folder)

        #edit theme
        theme = folder + r"""_copy\ppt\theme\%s""" % ("theme1.xml")
        new_theme=folder + r"""_copy\ppt\theme\%s""" % ("theme1.txt")
        os.rename(theme, new_theme)
        edit.edittext_(new_theme)  # edit acent text : -> _
        os.rename(new_theme, theme)
        #--------------------end edit-----------

        #xx, yy, zz, gg=edit.edit_name_txt2xml(title, folder)
        # xml2json.xml_2_json_getfromfile3()
        tree = ET.parse(xx)
        root = tree.getroot()
        # TRE_xml_rel
        TRE_xml_rel = ET.parse(link_rel_new)
        root_TRE_xml_rel = TRE_xml_rel.getroot()
        # xuly tree rels
        list_xmlres_Target = get_shape_obj_inrels_001(root_TRE_xml_rel)
        # lay img xml, shape xml after this
        shape_obj = get_shape_obj_img_shape_001(root)
        for numb,xx in enumerate(shape_obj):
            #print(xx.x,xx.y,xx.cx,xx.cy,xx.color,xx.prst,xx.img,xx.rId)
            for numbb,i in enumerate(list_xmlres_Target):
                #print('check target-id', i.Target, i.Id)
                if(i.Id==xx.rId):
                    xx.img='images'+i.Target[2:]
        #tree of THEME
        tree_oftheme = ET.parse(theme)
        root_oftheme = tree_oftheme.getroot()
        #end tree of theme
        print(tree)
        textngoai = get_text_ngoai(root,root_oftheme)
        texttrong = get_text_trong(root,root_oftheme)
        a = long_text(textngoai, texttrong,root_oftheme)
        finallytext = finallyoftext(a)
        # b=get_x_y__cxcy_text_trong(root)
        # fullobj_text(a,b)
        # checkRT_list(a)
        # get_text_trong(root)
        ngoai = get_x_y__cxcy_text_ngoai(root)
        trong = get_x_y__cxcy_text_trong(root)
        x_y = long_text_X_Y(ngoai, trong)
        print('check full obj textttt')
        fullobj_list = fullobj_text(listtext=finallytext, list_xy=x_y)
        xyz=checkfull_text_and_xy(fullobj_list)
    except Exception:
        traceback.print_exc()
    return xyz,shape_obj
def get_colorr(root_oftheme):
    #root_oftheme = tree_oftheme.getroot()
    list_oftheme = []
    list_of_obj_theme = []
    # for i in a:
    #     print("gia tri cuoi cung", i.text, i.x, i.y, i.cx, i.cy, i.color, i.size)
    for i in root_oftheme.findall('./a_themeElements/a_clrScheme/'):
        # print(i.tag,i.attrib,i.text)

        # print(i.tag, i.attrib, i.text)
        list_oftheme.append(i.tag)
    for numb, tagg in enumerate(list_oftheme):
        for i in root_oftheme.findall('./a_themeElements/a_clrScheme/' + str(tagg) + '/'):
            # print(i.tag,i.attrib,i.text)

            print('for in i, tag', numb, i.tag, i.attrib['val'], i.text)

            list_of_obj_theme.append(Create_obj_theme(i.tag, i.attrib['val']))
            # list_oftheme.append(i.tag)
    print('ket qua list_of obj')

    for numb, i in enumerate(list_of_obj_theme):
        print(numb, '   ', i.value_color, i.tag_color)
    for (f, b) in zip(list_of_obj_theme, list_oftheme):
        f.tag_color = b
    for i in list_of_obj_theme:
        print('finally ', i.tag_color, i.value_color)
    return list_of_obj_theme
if __name__ == "__main__":
    # a,b,c,d=edit.edit_name_xml2txt('slide5',r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json')
    # edit.edittext_(a)
    # edit.edit_name_txt2xml('slide5',r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json')
    # #xml2json.xml_2_json_getfromfile3()
    if False:
        tree= ET.parse(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\10slide_copy\ppt\slides\slide9.xml')
        root=tree.getroot()
        for i in root.findall('./p_cSld/p_spTree'):
            # print(i.tag,i.attrib,i.text)

            print(i.tag, i.attrib, i.text)
            #size
            for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../a_rPr'):
                print('value element ', x.text, x.tag, x.attrib['sz'],x.attrib)
            #color
            for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/...//a_solidFill/'):
                print('value element ', x.text, x.tag,x.attrib['val'])
                print(1)

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
        #a=lay_all_objtext_infilexml(folder=r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep\a1',title='slide5')
        a,b = lay_all_objtext_infilexml(folder=r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\10slide', title='slide9')
        tree_oftheme = ET.parse(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\10slide_copy\ppt\theme\theme1.xml')
        root_oftheme = tree_oftheme.getroot()
        list_oftheme=[]
        list_of_obj_theme=[]
        for i in a:
            print("gia tri cuoi cung",i.text,i.x,i.y,i.cx,i.cy,i.color,i.size)
        for i in root_oftheme.findall('./a_themeElements/a_clrScheme/'):
            # print(i.tag,i.attrib,i.text)

            #print(i.tag, i.attrib, i.text)
            list_oftheme.append(i.tag)
        for numb,tagg in enumerate(list_oftheme):
            for i in root_oftheme.findall('./a_themeElements/a_clrScheme/'+str(tagg)+'/'):
                # print(i.tag,i.attrib,i.text)

                print('for in i, tag',numb,i.tag, i.attrib['val'], i.text)

                list_of_obj_theme.append(Create_obj_theme(i.tag,i.attrib['val']))
                #list_oftheme.append(i.tag)
        print('ket qua list_of obj')

        for numb,i in enumerate(list_of_obj_theme):
            print(numb,'   ',i.value_color,i.tag_color)
        for (f, b) in zip(list_of_obj_theme, list_oftheme):
            f.tag_color=b
        for i in list_of_obj_theme:
            print('finally ',i.tag_color,i.value_color)
    #a,b=lay_all_objtext_infilexml(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\10slide','slide9')
    #get_color(root)

    tree= ET.parse(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\10slide_copy\ppt\slides\slide10.xml')
    root=tree.getroot()
    if False:
        for i in root.findall('./p_cSld/p_spTree'):
            # print(i.tag,i.attrib,i.text)

            # print(i.tag, i.attrib, i.text)
            for x in i.findall('./'):
                if x.tag == 'p_sp':
                    # get x y
                    aaa = Create_obj_shapeandimg('', '', '', '', '', '', '', '')
                    for j in x.findall('./p_spPr/a_xfrm/a_off'):
                        print(j.tag, j.attrib, j.attrib['x'], j.attrib['y'], j.text)
                        aaa.x = j.attrib['x']
                        aaa.y = j.attrib['y']
                        if (int(aaa.x) < 0):
                            aaa.x = 0
                        if (int(aaa.y) < 0):
                            aaa.y = 0

                    # get cx cy
                    for j in x.findall('./p_spPr/a_xfrm/a_ext'):
                        print(j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)

                        aaa.cx = j.attrib['cx']
                        aaa.cy = j.attrib['cy']
                        if (int(aaa.cx) < 0):
                            aaa.cx = 0
                        if (int(aaa.cy) < 0):
                            aaa.cy = 0

                    # for bf_j in i.findall('./p_sp/p_spPr/a_xfrm/a_off/.../a_ext'):
                    #     print(bf_j.attrib,bf_j.tag)
                    # for bf_j in i.findall('./p_sp/p_spPr/a_xfrm/a_off/.../a_ext'):
                    #     print(bf_j.attrib,bf_j.tag)
                    # for bf_j in i.findall('./p_sp/p_spPr/a_xfrm/a_off/.../a_ext'):
                    #     print(bf_j.attrib,bf_j.tag)
                    # for bf_j in i.findall('./p_sp/p_spPr/a_xfrm/a_off/.../a_ext'):
                    #     print(bf_j.attrib,bf_j.tag)
                    # for bf_j in i.findall('./p_sp/p_spPr/a_xfrm/a_off/.../a_ext'):
                    #     print(bf_j.attrib,bf_j.tag)
                    # for bf_j in i.findall('./p_sp/p_spPr/a_xfrm/a_off/.../a_ext'):
                    #     print(bf_j.attrib,bf_j.tag)

                    # get cx cy
                    # for j in i.findall('./p_sp/p_spPr/a_xfrm/a_ext'):
                    #     print(j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)
                    #
                    #     aaa.cx = j.attrib['cx']
                    #     aaa.cy = j.attrib['cy']
                    #     if (int(aaa.cx) < 0):
                    #         aaa.cx = 0
                    #     if (int(aaa.cy) < 0):
                    #         aaa.cy = 0
                    # # get prstGeom - shape - rec -line . . .
                    # for j in i.findall('./p_sp/p_spPr/a_prstGeom'):
                    #     print(j.tag, j.attrib, j.attrib['prst'], j.text)  # j.attrib['cx'], j.attrib['cy'],
                    #     aaa.prst = j.attrib['prst']
                    # # get solid_Fill
                    # for j in i.findall('./p_sp/p_spPr/a_solidFill/a_srgbClr'):
                    #     print(j.tag, j.attrib, j.attrib['val'],
                    #           j.text)  # j.attrib['val'] j.attrib['cx'], j.attrib['cy'],
                    #     aaa.color = j.attrib['val']
                    # list_shape.append(aaa)
                    # count = count + 1
                    # print('gia tri index count:', count)

                # if x.tag == 'p_pic':
                #     # get x y
                #     aaa = Create_obj_shapeandimg('', '', '', '', '', '', '', '')
                #     for j in i.findall('./p_pic/p_spPr/a_xfrm/a_off'):
                #         print(j.tag, j.attrib, j.attrib['x'], j.attrib['y'], j.text)
                #
                #         aaa.x = j.attrib['x']
                #         aaa.y = j.attrib['y']
                #         if (int(aaa.x) < 0):
                #             aaa.x = 0
                #         if (int(aaa.y) < 0):
                #             aaa.y = 0
                #     # get cx cy
                #     for j in i.findall('./p_pic/p_spPr/a_xfrm/a_ext'):
                #         print(j.tag, j.attrib, j.attrib['cx'], j.attrib['cy'], j.text)
                #
                #         aaa.cx = j.attrib['cx']
                #         aaa.cy = j.attrib['cy']
                #         if (int(aaa.cx) < 0):
                #             aaa.cx = 0
                #         if (int(aaa.cy) < 0):
                #             aaa.cy = 0
                #     # get prstGeom - shape - rec -line . . .
                #     for j in i.findall('./p_pic/p_spPr/a_prstGeom'):
                #         print(j.tag, j.attrib, j.attrib['prst'], j.text)  # j.attrib['cx'], j.attrib['cy'],
                #         aaa.prts = j.attrib['prst']
                #     # get Rid_img
                #     for j in i.findall('./p_pic/p_blipFill/a_blip'):  # /a_extLst
                #         print(j.tag, j.attrib['r_embed'], j.attrib,
                #               j.text)  # j.attrib['val'] j.attrib['cx'], j.attrib['cy'],
                #         aaa.rId = j.attrib['r_embed']
                #     list_shape.append(aaa)
                #     count = count + 1
                #     print('gia tri index count:', count)
        #print("gia tri cua count :", count)
    #check_list_shape(get_shape_obj_img_shape_001(root))
    a,b=lay_all_objtext_infilexml(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\10slide','slide9')
    #for i,data in enumerate(b):

    for j,dataa in enumerate(b):
        if dataa.prst =='rect':
            print(' none dataa.img',dataa.img,'dataa.prst',dataa.prst,dataa.color)
        else:
            if dataa.img[-3:]== 'png':
                print('deptrai')
            print(j,dataa.img,dataa.rId,dataa.prst,dataa.x,dataa.y,dataa.cx,dataa.cy,dataa.color)
            dataa.x, dataa.y, dataa.cx, dataa.cy = nhap5_usedict_addobj.convert_donvi_text_xycxcy_2_h5p(dataa.x,dataa.y,dataa.cx,dataa.cy)
            print('tesst dataa',dataa.x,dataa.y,dataa.cx,dataa.cy)
# for numb,xx in enumerate(shape_obj):
#         #print(xx.x,xx.y,xx.cx,xx.cy,xx.color,xx.prst,xx.img,xx.rId)
#         for numbb,i in enumerate(list_xmlres_Target):
#             #print('check target-id', i.Target, i.Id)
#             if(i.Id==xx.rId):
#                 xx.img='images'+i.Target[2:]
    bb = [i for i, list_shape in enumerate(b) if not b]
    nhap5_usedict_addobj.dithang_shape_object_ver001(b,bb)
