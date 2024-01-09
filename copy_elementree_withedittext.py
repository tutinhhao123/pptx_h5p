from builtins import enumerate

import js2py
import xml.etree.ElementTree as ET
import os
import xmltodict
import edit_colon2_ as edit
import json
import xml.dom.minidom
#import xml2jsonn as xml2json
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
                # , zz.text, zz.tag
                print("value zz", zz.text, zz.tag, zz.attrib)
                if (int(zz.attrib['y']) < 0):
                    a.y = '0'
                else:
                    a.y = zz.attrib['y']
                if (int(zz.attrib['x']) < 0):
                    a.x = '0'
                else:
                    a.x = zz.attrib['x']

                # a.x = zz.attrib['x']
            for zz in x.findall('.p_spPr/a_xfrm/a_ext'):
                if (int(zz.attrib['cy']) < 0):
                    a.cy = '0'
                else:
                    a.cy = zz.attrib['cy']
                if (int(zz.attrib['cx']) < 0):
                    a.cx = '0'
                else:
                    a.cx = zz.attrib['cx']
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
                if(int(zz.attrib['y'])<0):
                    a.y = '0'
                else:
                    a.y = zz.attrib['y']
                if (int(zz.attrib['x']) < 0):
                    a.x = '0'
                else:
                    a.x = zz.attrib['x']

                #a.x = zz.attrib['x']
            for zz in x.findall('.p_spPr/a_xfrm/a_ext'):
                if (int(zz.attrib['cy']) < 0):
                    a.cy = '0'
                else:
                    a.cy = zz.attrib['cy']
                if (int(zz.attrib['cx']) < 0):
                    a.cx = '0'
                else:
                    a.cx = zz.attrib['cx']
                # a.cx = zz.attrib['cx']
                # a.cy = zz.attrib['cy']


            list.append(a)
    checkRT_list(list)
    return list
def checkRT_list(aa):
    for i in aa:
        print('check rt_list',i.x,i.y )
def checkRT_img_XMLRELS_list(aa):
    for i in aa:
        print('check target-id',i.Target[2:],i.Id )
        print('images'+i.Target[2:])
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
def get_shape_obj_img_shape_001(root):
    list_shape = []
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

    return list_shape
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
def lay_all_objtext_infilexml(folder,title):
    aa, bb, cc, dd,link_rel = edit.edit_name_xml2txt(title, folder)
    edit.edittext_(aa)
    edit.edittext_(link_rel)

    xx, yy, zz, gg,link_rel_new=edit.edit_name_txt2xml(title, folder)
    # xml2json.xml_2_json_getfromfile3()
    tree = ET.parse(xx)
    root = tree.getroot()
    #TRE_xml_rel
    TRE_xml_rel=ET.parse(link_rel)
    root_TRE_xml_rel=TRE_xml_rel.getroot()
    #xuly tree rels

    list_xmlres_Target = get_shape_obj_inrels_001(root_TRE_xml_rel)
    #lay img xml, shape xml after this
    shape_obj=get_shape_obj_img_shape_001(root)

    for numb,xx in enumerate(shape_obj):
        #print(xx.x,xx.y,xx.cx,xx.cy,xx.color,xx.prst,xx.img,xx.rId)
        for i in list_xmlres_Target:
            #print('check target-id', i.Target, i.Id)
            if(i.Id==xx.rId):
                xx.img='images'+i.Target[2:]

    #code . . .

    #end lay ig xml shape xml after ths
    #_--------end tree rel
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
    return xyz,shape_obj

if __name__ == "__main__":
    # a,b,c,d=edit.edit_name_xml2txt('slide5',r'C:\Users\Admin\anaconda3\envs\h5p_11\xml2json')
    # edit.edittext_(a)
    # edit.edit_name_txt2xml('slide5',r'C:\Users\Admin\anaconda3\envs\h5p_11\xml2json')
    # #xml2json.xml_2_json_getfromfile3()
    tree_xml_rel= ET.parse(r'C:\Users\ACER\anaconda3\envs\h5p\pptx2h5p_1_2_deep_copy_addimages-20230319T030947Z-001\pptx2h5p_1_2_deep_copy_addimages\du_an_test01_copy\ppt\slides\_rels\slide9.xml.rels')
    root_tree_xml_rel=tree_xml_rel.getroot()
    print(tree_xml_rel)
    # tree= ET.parse(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\a1_copy\ppt\slides\slide9.xml')
    # root=tree.getroot()
    # print(root)
    # print(get_shape_obj_img_shape_001(root))
    # a=get_shape_obj_img_shape_001(root)
    #
    # for numb,xx in enumerate( a):
    #     print(xx.x,xx.y,xx.cx,xx.cy,xx.color,xx.prst,xx.img,xx.rId)
        #for j in enumerate (x):
    get_shape_obj_inrels_001(root_tree_xml_rel)

    # for i in root.findall('.//'):
    #     # print(i.tag,i.attrib,i.text)
    #     # if(i.tag == 'Relationship' or i.attrib == 'Relationship'):
    #     #     print('tag ss')
    #     print( i.attrib['Target'],i.attrib)
    # getshape=get_shape_obj_inrels_001(root)
    #
    # for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../.../.../...'):
    #     print(i.tag, i.attrib, i.text)
    #
    # for x in root.findall('./p_cSld/p_spTree//p_pic'):
    #         print(x.tag, x.attrib)#, x.text)

            #print(x.tag, x.attrib, x.text)
    # print('end')
            # for j in x.findall('.//'):
        #     print(j.tag,j.attrib)
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



    # a=lay_all_objtext_infilexml(folder=r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep_copy_addimages\pptx2h5p_1_2_deep_copy_addimages\a1',title='slide5')
    # for i in a:
    #     print("gia tri cuoi cung",i.text,i.x,i.y,i.cx,i.cy)