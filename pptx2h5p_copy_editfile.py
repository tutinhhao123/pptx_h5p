import json
import os
import sys
import uuid
from copy import deepcopy
from zipfile import ZipFile
import shutil
import patoolib
from natsort import natsorted
from win32com import client
from get_image_size import get_image_size
path_file_ptpx_rar=r""
VERSION = "1.2"
YEAR = "2021"
target_ratio = 2  # target aspect ratio for slides in h5p
from imread_XML_infolder import load_xml_from_folder
from xml2json import xml2json
#from edit_string_pptx2rar import edit_name_pptx2rar
#xml2json.json_object
#import xml2json.xml2json as xml2json
#import xml2json.elementree_withedittext
class Create_obj_p():          # leave this empty
    def __init__(self):   # constructor function using self
        self.text = None  # variable using self.
        self.color = None  # variable using self

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
#     old_name = str(title) + r"C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p-1.2_dêp\a - Copy.pptx"
#     new_name = str(folder) + r"C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p-1.2_dêp\a - Copy.rar"
#     # Renaming the file
#     os.rename(old_name, new_name)

# method tự động sinh text từ pptx to h5p
def converttext_h5p_mini(text,color):
    return r"""<span><span><span><span><span style=\"color:%s\"><span><span><span>%s</span></span></span></span> """%("color","text")
def create_arrObj(text1,color1):
    Array = []
    obj_1 = Create_obj_p()
    obj_1.text = text1
    obj_1.color = color1
    Array.append(obj_1)
    return Array

def cpmverttext_h5p(obj_text_fromxml):
    #read xml -> add obj to array_obj
    #load_xml_from_folder()
    html=''
    for i in obj_text_fromxml:
        a = converttext_h5p_mini(i.text, i.color)
        html = html+str(a)
    return r"""<p style=\"text-align:left\"><span style=\"font-size:1.5em\">%s</span></p>\n"""%(html)

def add_to_json(newfile, image_folder, images, title, array_Obj,linkarr_xml,titleslide):
    exclude_files = ['content/content.json', 'h5p.json']
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(os.path.abspath(__file__))
    template = os.path.join(basedir, "template.h5p")

    img_width, img_height = get_image_size(os.path.join(image_folder, images[0]))

    with ZipFile(template, "r") as zin:
        with ZipFile(newfile, "w") as zout:
            # copy all other files
            for item in zin.infolist():
                if item.filename not in exclude_files:
                    zout.writestr(item, zin.read(item.filename))

            # add image filenames to content.json
            with zin.open("content/content.json") as fp:
                content = json.load(fp)
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
                #god bles you ♥
                for i, link_xml_file in enumerate(linkarr_xml):
                    #xml after cv 2 json- lấy hết obj from từng xml
                    obj_xml_cv2json,obj_rutgon = xml2json.xml_2_json_getfromfile3(link_xml_file)
                    print("gia tri cua array_obj_xml",obj_xml_cv2json)

                    #get element to add param text to json

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
                    # set width & height
                    params["file"]["width"] = img_width
                    params["file"]["height"] = img_height


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

if __name__ == "__main__":
    try:
        # Manifest
        print("Powerpoint to h5p Converter.")
        print(f"Version: {VERSION}")
        print(f"Martin Lehmann, {YEAR}")
        print("Licence: BSD-2-Clause")
        print("Source code: https://github.com/MM-Lehmann/pptx2h5p")
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
        #load link xml
        XML,linkXML,c = load_xml_from_folder(os.path.join(folderr, ten_folder_rar))
        print('load xml from folder', XML)
        # add_to_json(
        #     newfilename, image_folder, images, title, array_Obj=XML,linkarr_xml=linkXML,titleslide=c
        # )
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
