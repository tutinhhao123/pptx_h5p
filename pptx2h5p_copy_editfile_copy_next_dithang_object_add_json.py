import json
import os
import sys
import uuid
from copy import deepcopy
from zipfile import ZipFile
from natsort import natsorted
from win32com import client
from get_image_size import get_image_size
path_file_ptpx_rar=r""
VERSION = "1.2"
YEAR = "2021"
target_ratio = 2  # target aspect ratio for slides in h5p
#import xml2json.json_object
#import xml2json.elementree_withedittext
import elementree_withedittext
import xml.etree.ElementTree as ET
import edit_colon2_ as edit
from edit_string_pptx2rar import edit_name_pptx2rar as edit
#import xml2json.elementree_withedittext as elementree_withedittext
from imread_XML_infolder import load_xml_from_folder
from edit_string_pptx2rar import edit_name_pptx2rar
#import xml2json.json_object
#import xml2json.elementree_withedittext
#import xml2json.elementree_withedittext as elementree_withedittext

class Create_obj_p():          # leave this empty
    def __init__(self):   # constructor function using self
        self.text = None  # variable using self.
        self.color = None  # variable using self

def dithangobject(arrayobj):
    #r"""\%s""" % ("")
    #x,y,width,height,text,subContentId
    #x,y,width,height= number -> nothing
    #text,subcontent is string -> ""
    h5ptext = r"""{
                "x": %s,
                "y": %s,
                "width": %s,
                "height": %s,
                "action": {
                  "library": "H5P.AdvancedText 1.1",
                  "params": {
                    "text": "%s"      
                    },
                  "subContentId": "%s",
                  "metadata": {
                    "contentType": "Text",
                    "license": "U",
                    "title": "Untitled Text",
                    "authors": [],
                    "changes": []
                  }
                },
                "alwaysDisplayComments": false,
                "backgroundOpacity": 0,
                "displayAsButton": false,
                "buttonSize": "big",
                "goToSlideType": "specified",
                "invisible": false,
                "solution": ""
              }"""
    listobj_text=[]
    list_none=[]
    return listobj_text,list_none
def ppt2png(file):
    try:
        powerpoint = client.Dispatch("Powerpoint.Application")
    except Exception as e:
        print("Powerpoint could not be opened", file=sys.stderr)
        raise e
    try: # look if an active presentation is open
        assert powerpoint.ActivePresentation is not None
        QUIT = False # don't quit powerpoint later
    except:
        QUIT = True # quit powerpoint later
    ppt = powerpoint.Presentations.Open(file)
    ppt.Export(file, "PNG")
    ppt.Close()
    if QUIT: # quit only if required
        powerpoint.Quit()
####lưu tâm mốt sửa sau!!!! ôi buồn quá!
# def edit_name_pptx2rar(title,folder):
#     original = r'C:\Users\Ron\Desktop\Test_1\products.csv'
#     target = r'C:\Users\Ron\Desktop\Test_2\products.csv'
#     shutil.copyfile(original, target)
#     old_name = str(title) + r"C:\Users\Admin\anaconda3\envs\h5p_11\pptx2h5p-1.2_dêp\a - Copy.pptx"
#     new_name = str(folder) + r"C:\Users\Admin\anaconda3\envs\h5p_11\pptx2h5p-1.2_dêp\a - Copy.rar"
#     # Renaming the file
#     os.rename(old_name, new_name)

# method tự động sinh text từ pptx to h5p
def converttext_h5p_mini(text):
    return r"""<span><span><span><span><span style=\"color:#000000\"><span><span><span>%s</span></span></span></span> """%(text)
def create_arrObj(text1,color1):
    Array = []
    obj_1 = Create_obj_p()
    obj_1.text = text1
    obj_1.color = color1
    Array.append(obj_1)
    return Array

def cpmverttext_h5p(text):
    #read xml -> add obj to array_obj
    #load_xml_from_folder()
    html=''

    a = converttext_h5p_mini(text)
    html = html+str(a)
    return r"""<p style=\"text-align:left\"><span style=\"font-size:1.5em\">%s</span></p>\n"""%(html)
def convert_donvi_text_xycxcy_2_h5p(x,y,cx,cy):
    # p: sldSz
    # type = "screen4x3"
    # cy = "6858000"
    # cx = "9144000" /
    x=str(100*int(x)/9144000)

    y=str(100*int(y)/6858000)
    cx = str(100 * int(cx) / 9144000)
    cy = str(100 * int(cy) / 6858000)
    print('new value off donvitext x y cx cy =',x,y,cx,cy)
    return x,y,cx,cy
#def add_to_json(newfile, image_folder, images, title,arrayobj, folder_xml,linkarr_xml,title_xml):
def add_to_json(newfile, image_folder, images, title, arrayobj):

    exclude_files = ['content/content.json', 'h5p.json']
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(os.path.abspath(__file__))
    template = os.path.join(basedir, "template_new.h5p")

    img_width, img_height = get_image_size(os.path.join(image_folder, images[0]))

    with ZipFile(template, "r") as zin:
        with ZipFile(newfile, "w") as zout:
            # copy all other files
            for item in zin.infolist():
                print("item.filename",item.filename)
                if item.filename not in exclude_files:
                    print(zin.read(item.filename))
                    zout.writestr(item, zin.read(item.filename))

            # add image filenames to content.json
            with zin.open("content/content.json") as fp:
                content = json.load(fp)
                #----------------------
                #print('value content is:',content)
                #print('value content dump:',json.dumps(content, indent=1))
                # ----------------------
                # for ii in content:
                #     print('value ii',ii)
                #print('conten',content)
                print(f"adding images from {image_folder}: {images}")
                for i, image in enumerate(images):
                    slides = content["presentation"]["slides"]
                    # clone first element
                    if i > 0 and len(slides) <= i:
                        slides.append(deepcopy(slides[0]))
                    elements = slides[i]["elements"][0]
                    params = elements["action"]["params"]
                    # add new image filename
                    params["file"]["path"] = "images/" + image
                    # add random uuid
                    elements["action"]["subContentId"] = str(uuid.uuid4())
                    #set width & height
                    params["file"]["width"] = img_width
                    params["file"]["height"] = img_height
                    ratio = img_width / img_height
                    if ratio > target_ratio:  # wider, need to shrink y
                        elements["y"] = 100 * (1 - target_ratio / ratio) / 2
                        elements["height"] = 100 * target_ratio / ratio
                    elif ratio < target_ratio:  # higher, need to shrink x
                        elements["x"] = 100 * (1 - ratio / target_ratio) / 2
                        elements["width"] = 100 * ratio / target_ratio
                # #god bles you ♥
                            # for i, x in enumerate(listt):
                            #
                            #     print(i, x)
                            #     print('#-------------------------text obj.x.y-------------------')
                            #
                            #     for j in x:
                            #         print(i, j, ',', j.text, j.x, j.y)
                            #         # for z in j:
                            #         #  print('x:',z.text)
                            #
                            #     print('#-------------------------end text obj.x.y-------------------')
                #slides = content["presentation"]["slides"]

                # for x in slides:
                #     print('x in slides',x)
                # print('print slide 0:',slides[0])
                # print('print slide 1:',slides[1])
                # #god bles you ♥
                a = [i for i, x in enumerate(arrayobj) if not x]
                for x in a:
                    print('values arrayobj none:', x)
                for i, arr_obj in enumerate(arrayobj):
                    if i not in a:

                        # print(i)
                        # for j,dataa in enumerate(arr_obj):
                        #     print(j,'text-cx-cy',dataa.text,dataa.cx,dataa.cy)
                        #xml after cv 2 json- lấy hết obj from từng xml
                        #obj_xml_cv2json,obj_rutgon = xml2json.xml_2_json_getfromfile3(link_xml_file)
                        print('slide', i)
                        print("gia tri cua arrayobj",i,arr_obj)
                        print('#-------------------------text obj.x.y-------------------')
                        #get element to add param text to json
                        #edit element slides ♥ god bles u
                        slides = content["presentation"]["slides"]
                        # clone first element
                        # for x in slides:
                        #     print('x in slides',x)
                        #print()
                        #tính lại x,y, cx,cy
                        #viet method
                        #convert_donvi_text_xycxcy_2_h5p()
                        print('#-------------------------end text obj.x.y-------------------')


                        #if not arr_obj:
                        #print(i, 'i none')
                        #continue
                        #else:


                        for j,data in enumerate(arr_obj):
                            #add element text cho tung slide

                            if j > 0 and len(arr_obj) <= j :
                                #xx = xx + 1
                                slides[i]['elements'].append(deepcopy(slides[0]['elements'][1]))
                            #tính lại x,y, cx,cy
                                    # viet method
                            if(1==1):
                                #----------------------------------------- phai xu ly duoc list index
                                data.x,data.y,data.cx,data.cy=convert_donvi_text_xycxcy_2_h5p(data.x,data.y,data.cx,data.cy)
                                elements = slides[i]["elements"][j+1]
                                slides[i]["elements"][j+1]['x']=data.x#['x']
                                slides[i]["elements"][j+1]['y'] = data.y#['y']
                                slides[i]["elements"][j+1]['width'] = data.cx#['cx']
                                slides[i]["elements"][j+1]['height'] = data.cy#['cy']

                                params = elements["action"]["params"]
                                # add new image filename
                                #---------------------------
                                params["text"] = cpmverttext_h5p(cpmverttext_h5p(data.text))
                                #---------------------------
                                # add random uuid
                                elements["action"]["subContentId"] = str(uuid.uuid4())
                                # set width & height
                                # params["file"]["width"] = img_width
                                # params["file"]["height"] = img_height
                        print('#-------------------------end text obj.x.y-------------------')
                            #else:
                            #xx=xx-1
                        #xx=xx+1
                    #

                # for i, arr_obj in enumerate(arrayobj):
                #     for j, data in enumerate(arr_obj):
                #
                #         # tính lại x,y, cx,cy
                #         # viet method
                #         if(1==1 ):
                #             data.x,data.y,data.cx,data.cy=convert_donvi_text_xycxcy_2_h5p(data.x,data.y,data.cx,data.cy)
                #             elements = slides[i]["elements"][j]
                #             slides[i]["elements"][j]['x']=data.x#['x']
                #             slides[i]["elements"][j]['y'] = data.y#['y']
                #             slides[i]["elements"][j]['width'] = data.cx#['cx']
                #             slides[i]["elements"][j]['height'] = data.cy#['cy']
                #
                #             params = elements["action"]["params"]
                #             # add new image filename
                #             #---------------------------
                #             params["text"] = cpmverttext_h5p(cpmverttext_h5p(data.text))
                #             #---------------------------
                #             # add random uuid
                #             elements["action"]["subContentId"] = str(uuid.uuid4())
                #             # set width & height
                #             # params["file"]["width"] = img_width
                #             # params["file"]["height"] = img_height
                # print('#-------------------------end text obj.x.y-------------------')
                for i, image in enumerate(images):
                    slides = content["presentation"]["slides"]
                    # clone first element
                    # if i > 0 and len(slides) <= i:
                    #     slides.append(deepcopy(slides[0]))
                    elements = slides[i]["elements"][0]
                    params = elements["action"]["params"]
                    # add new image filename
                    params["file"]["path"] = "images/" + image
                    # add random uuid
                    elements["action"]["subContentId"] = str(uuid.uuid4())
                    #set width & height
                    params["file"]["width"] = img_width
                    params["file"]["height"] = img_height
                    ratio = img_width / img_height
                    if ratio > target_ratio:  # wider, need to shrink y
                        elements["y"] = 0#100 * (1 - target_ratio / ratio) / 2
                        elements["height"] = 100# * target_ratio / ratio
                    elif ratio < target_ratio:  # higher, need to shrink x
                        elements["x"] = 0#100 * (1 - ratio / target_ratio) / 2
                        elements["width"] = 100 #* ratio / target_ratio
            # save file
            zout.writestr("content/content.json", json.dumps(content))

            # add image files to tip
            for image in images:
                zout.write(os.path.join(image_folder, image), "content/images/" + image)

            # change presentation title
            with zin.open("h5p.json", "r") as h5p:
                content = json.load(h5p)
                content["title"] = title
            zout.writestr("h5p.json", json.dumps(content))

# def extra_rarfile(namefile,path):
#     patoolib.extract_archive(namefile, outdir=path)
def lay_all_objtext_infilexml(folder,title):
    aa, bb, cc, dd = edit.edit_name_xml2txt(title, folder)
    edit.edittext_(aa)
    xx, yy, zz, gg=edit.edit_name_txt2xml(title, folder)
    # xml2json.xml_2_json_getfromfile3()
    tree = ET.parse(xx)
    root = tree.getroot()
    print(tree)
    textngoai = elementree_withedittext.get_text_ngoai(root)
    texttrong = elementree_withedittext.get_text_trong(root)
    a = elementree_withedittext.long_text(textngoai, texttrong)

    finallytext = elementree_withedittext.finallyoftext(a)
    # b=get_x_y__cxcy_text_trong(root)
    # fullobj_text(a,b)
    # checkRT_list(a)
    # get_text_trong(root)
    ngoai = elementree_withedittext.get_x_y__cxcy_text_ngoai(root)
    trong = elementree_withedittext.get_x_y__cxcy_text_trong(root)
    x_y = elementree_withedittext.long_text_X_Y(ngoai, trong)
    fullobj_list = elementree_withedittext.fullobj_text(listtext=finallytext, list_xy=x_y)
    #xyz=checkfull_text_and_xy(fullobj_list)
    for i in fullobj_list:
        i.x,i.y,i.cx,i.cy=convert_donvi_text_xycxcy_2_h5p(i.x,i.y,i.cx,i.cy)
    return elementree_withedittext.checkfull_text_and_xy(fullobj_list)
if __name__ == "__main__":
    try:
        # Manifest
        print("Powerpoint to h5p Converter.")
        print(f"Version: {VERSION}")
        print(f"Martin Lehmann, {YEAR}")
        print("Licence: BSD-2-Clause")
        print("Source code: https://github.com/MM-Lehmann/pptx2h5p")

        print("develop by tran tuan anh")
        print('he is superman')
        if len(sys.argv) != 2:
            print("Usage : python pptx2h5p.py [file]", file=sys.stderr)
            sys.exit(-1)

        # extract metadata
        filepath = os.path.abspath(sys.argv[1])
        print(f"extracting images from {filepath}.")
        folder = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        title = os.path.splitext(filename)[0]
        if not os.path.exists(filepath):
            print("No such file!", file=sys.stderr)
            sys.exit(-1)

        # extract images
        ppt2png(filepath)
        #sua images folder
        image_folder = os.path.join(folder, title)
        images = natsorted(os.listdir(image_folder))

        # compile .hp5 file
        newfilename = os.path.splitext(filepath)[0] + ".h5p"
        print("building new "+str(newfilename)+" file")
        print(r" \n gia tri cua title la:"+str(title)+r"\n gia tri cua folder"+str(folder))
        #load xml from da giai nen
        newname_fodler_rar, folderr, new_title_rar, ten_folder_rar = edit_name_pptx2rar(title,folder)
        print('newname_fodler_rar:',newname_fodler_rar)
        print('new_title_rar:', new_title_rar)
        print('folderr:', folderr)
        print('ten_folder_rar:', ten_folder_rar)
        #--------------------------------------
        #xử lý lại extra file sau
        #extra_rarfile(new_title_rar, folderr)
        #--------------------------------------
        print('extra rar file')
        #load_xml_from_folder(os.path.splitext(filepath)[0])
        # #load link xml
        name_XMLa, linkfile_xml, xxxxx = load_xml_from_folder(
            os.path.join(folder, title))
        print("loadxml from folder", name_XMLa)
        print("link file xml", linkfile_xml)
        print("gia tri:", linkfile_xml[1])
        print('#------------------------------------')
        print("xxxxx", xxxxx)
        print('#------------------------------------')
        listfiletext=[]

        for i in xxxxx:

             listfiletext.append(elementree_withedittext.lay_all_objtext_infilexml(os.path.join(folder, title), i))

        # print('#---------listfiletext---------------------------')
        # #print(listfiletext)
        # # for i in listfiletext:  # i['text'],
        # #     print("gia tri cuoi cung", i['x'], i['y'], i['cx'], i['cy'])
        # for i,data in enumerate (listfiletext):
        #     print (i)
        #     for j, dataa in enumerate(data):
        #       print(j,"x-y-cx-cy",dataa.text, dataa.x)#, j['y'], j['cx'], j['cy'], j['text'])


        # add_to_json(
        # #     newfilename, image_folder, images, title, array_Obj=XML,linkarr_xml=linkXML,titleslide=c
        # # )
        add_to_json(
                newfilename, image_folder, images, title, arrayobj =listfiletext
            )
        print("Converting successfully finished.")
        input(
            "Press Enter to delete temporary image (export) folder and close this window."
        )
        for image in images:
            os.remove(os.path.join(image_folder, image))
        os.rmdir(image_folder)

    except Exception as e:
        print(e, file=sys.stderr)
        input("An error has occured. Press Enter to close this window.")
