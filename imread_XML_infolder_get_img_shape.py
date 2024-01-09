import cv2
import os

from natsort import natsorted,natsort_keygen

#from xml2jsonn import xml2json
import xmltodict
#import xml2json
import json
#file = xmltodict.parse(r'C:\Users\Admin\anaconda3\envs\h5p_11\pptx2h5p-1.2_dêp\slide1.xml')
from xml.dom import minidom
# parse file items.xml
import codecs
import xml.etree.ElementTree as ET
from copy_elementree_withedittext import lay_all_objtext_infilexml
types_of_encoding = ["utf8", "cp1252"]
from xml.dom.minidom import parse, parseString
#import xml2json.xml2json as xml2json
import edit_colon2_
#method load all link xml in folder

def load_xml_from_folder(folder):
    XMLS = []
    link=[]
    newfoler=[]
    listtitle=[]
    #a=r
    for filename in os.listdir(folder +str( """_copy\ppt""") + str("""\slides""")):
        #print(filename)
        #xml = cv2.imread(os.path.join(folder + """\ppt""" + """\slides""", filename))
        #xml = cv2.imread(os.path.join(folder, filename))
        xml = filename
        #xml = natsorted(os.listdir(folder + """\ppt""" + """\slides"""))

        if xml is not None:
            XMLS.append(xml)
            #link.append(folder + """\ppt""" + """\slides\%s"""%(xml))
            listtitle.append(xml[:-4])
        #XMLS.sort(key=lambda x: float(x.strip('slide')))
        #XMLS.remove('_rels')
        #XMLS.append(xml)
    #natsort_key = natsort_keygen()
    #print(natsort_key)
    #newxml=natsorted(XMLS)
    newxml = list.copy(natsorted(XMLS))
    newlisttitle= list.copy(natsorted(listtitle))
    for filename in XMLS:
        #link.append(os.listdir(folder str( """\ppt""") + str("""\slides"""+r"""\%s"""%(filename))) )
        link.append(os.path.join(folder +str( """_copy\ppt""") + str("""\slides"""), filename))
        #newfoler.append()
    newlink = list.copy(natsorted(link))

    #xmlsss=newxml[]
    #link_xml_cuapptx = folder + """\ppt""" + """\slides"""
    return newxml[1:],newlink[1:],newlisttitle[1:]#,newfoler
def load_xml_from_folder_togetims_shape(folder):
    XMLS = []
    link=[]
    newfoler=[]
    listtitle=[]
    #a=r
    for filename in os.listdir(folder +str( """_copy\ppt""") + str("""\slides""")):
        #print(filename)
        #xml = cv2.imread(os.path.join(folder + """\ppt""" + """\slides""", filename))
        #xml = cv2.imread(os.path.join(folder, filename))
        xml = filename
        #xml = natsorted(os.listdir(folder + """\ppt""" + """\slides"""))

        if xml is not None:
            XMLS.append(xml)
            #link.append(folder + """\ppt""" + """\slides\%s"""%(xml))
            listtitle.append(xml[:-4])

        #XMLS.sort(key=lambda x: float(x.strip('slide')))

        #XMLS.remove('_rels')
        #XMLS.append(xml)
    #natsort_key = natsort_keygen()
    #print(natsort_key)
    #newxml=natsorted(XMLS)
    newxml = list.copy(natsorted(XMLS))
    newlisttitle= list.copy(natsorted(listtitle))
    print('gia tri new list tilte',newlisttitle)
    i=0
    for filename in XMLS:
        #link.append(os.listdir(folder str( """\ppt""") + str("""\slides"""+r"""\%s"""%(filename))) )
        link.append(os.path.join(folder +str( """_copy\ppt""") + str("""\slides"""), filename))
        #newfoler.append()

    newlink = list.copy(natsorted(link))
    _rel=newlink[0]
    # print('gia tri link _rel:',_rel)
    #xmlsss=newxml[]
    #link_xml_cuapptx = folder + """\ppt""" + """\slides"""
    return newxml[1:],newlink[1:],newlisttitle[1:],_rel#,newfoler

if __name__ == "__main__":
    # string = r"C:\Users\Admin\anaconda3\envs\h5p_11\pptx2h5p_1_2_deep\a1"
    # #XMLS = load_xml_from_folder(string)
    # #print("vaalues cuối:",XMLS)

    name_XMLa,linkfile_xml,xxxxx,_rel = load_xml_from_folder_togetims_shape(r"C:\Users\ACER\anaconda3\envs\h5p\pptx2h5p_1_2_deep_copy_addimages-20230319T030947Z-001\pptx2h5p_1_2_deep_copy_addimages\du_an_test01")
    print("loadxml from folder",name_XMLa)
    print("link file xml",linkfile_xml)
    print("gia tri:", linkfile_xml[1])
    print("xxxxx",xxxxx)
    for i in xxxxx:
        lay_all_objtext_infilexml(r"C:\Users\ACER\anaconda3\envs\h5p\pptx2h5p_1_2_deep_copy_addimages-20230319T030947Z-001\pptx2h5p_1_2_deep_copy_addimages\du_an_test01",i)

    # # a,b=xml2json.xml_2_json_getfromfile3(linkfile_xml[1])
    # print(b)
    # document = parse(linkfile_xml[0])
    # print(document)
    # for file in linkfile_xml:
    #     #print(xml2json.xml_2_json_getfromfile3(file))
    #     print('file')

    # a=xml2json.xml_2_json_getfromfile3(os.fsencode(r'C:\\Users\\Admin\\anaconda3\\envs\\h5p_11\\pptx2h5p-1.2_deep\\a1\\ppt\\slides\\slide1.xml'))
    # print("gia tri cua a",a)
    # a,b = xml2json.xml_2_json_getfromfile3(
    #     linkfile_xml)
    # print("gia tri cua a", b)

    # # for i in range(len(a)):
    #     a[i] = a[i].rstrip()
    # print(a)
    # for i, xml_file in enumerate(linkfile_xml):
    #     # xml after cv 2 json
    #     obj_xml_cv2json,newlinh= xml2json.xml_2_json_getfromfile3(xml_file)
    # print(obj_xml_cv2json,newlinh)