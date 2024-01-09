import xmltodict
import json
import lxml
import js2py
import xml.etree.ElementTree as ET
import os
import xmltodict
import edit_colon2_ as edit
import json
import xml.dom.minidom
import xml2json as xml2json
import pyutil
with open(r"C:\Users\Admin\anaconda3\envs\backuptamthoi3\pptx2h5p_1_2_deep_copy_addimages\10slide_copy\ppt\slides\slide1.xml", "rb") as file:
     document = xmltodict.parse(file)
     print(document.keys())
# with open(r"C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p-1.2_deep\a1\ppt\slides\slide1.xml", "rb") as file:
#      document = xmltodict.parse(file, process_namespaces=True)
#      print(document.keys())
print(json.dumps(document, indent=4))
a,b,c,d=edit.edit_name_xml2txt('slide5',r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json')
edit.edittext_(a)
edit.edit_name_txt2xml('slide5',r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json')
#xml2json.xml_2_json_getfromfile3()
tree= ET.parse(r'C:\Users\Admin\anaconda3\envs\backup_1\xml2json\slide5.xml')
root=tree.getroot()
j = 0
list = []


text = ""

#------------------------------------------
for i in root.findall('./p_cSld/p_spTree'):
     # print(i.tag,i.attrib,i.text)

     print(i.tag, i.attrib, i.text)
     for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../.../.../...'):
          # print(x.tag, x.attrib, x.text)
           print('value element ', x.text, x.tag, x.attrib)
           for zz in x.findall('.p_spPr/a_xfrm/'):
                print("value zz",zz.text, zz.tag, zz.attrib)
#---------------------------------------------------------
# for i in root.findall('./p_cSld/p_spTree/p_grpSp'):
#     # print(i.tag,i.attrib,i.text)
#
#     # print(i.tag, i.attrib, i.text)
#     for x in i.findall('./p_sp/p_txBody/a_p/a_r/a_t/.../.../.../...'):
#         #print(x.tag, x.attrib, x.text)
#        #print('value element ', x.text, x.tag, x.attrib)
#         #p_spPr/a_xfrm/a_off
#        for zz in x.findall('.p_spPr/a_xfrm/'):
#             print("value zz",zz.text, zz.tag, zz.attrib)
#------------------------------------------------------
          #
          # j += 1
          # if (x.text != " " and x.text != ""):
          #      text = text + x.text
          # print(j)
          # # list.append(Create_obj_p(text, 3))
          # print(x.tag,x.attrib)
          # #elm = x.findall('..')
          # #elm.getparent()
          # for z in x.findall('...'):
          #      print('value element ',z.text,z.tag,z.attrib)
     #print("textlay duoc", text)