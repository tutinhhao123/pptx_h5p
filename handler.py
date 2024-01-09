import nhap4
from nhap4 import traveobj_slide,convert_donvi_text_xycxcy_2_h5p
import json
def dithangobject_ver001(arrayobj,arr_none):
    #r"""\%s""" % ("")
    #x,y,width,height,text,subContentId
    #x,y,width,height= number -> nothing
    #text,subcontent is string -> ""
    print('#-------------di thang obj-------------------')
    listobj_textfield=[]
    listobj_text_slide=[]
    list_none=arr_none
    for i,x in enumerate(arrayobj):

        if not x:
            print('slide ',i,'None')
            listobj_text_slide.append('')
        else:
            print('slide truoc add', i)
            for j,z in enumerate(x):
                #cpmverttext_h5p

                z.x, z.y, z.cx, z.cy = convert_donvi_text_xycxcy_2_h5p(z.x, z.y, z.cx, z.cy)

                h5ptext = {
                                "x": z.x,
                                "y": z.y,
                                "width": z.cx,
                                "height": z.cy,
                                "action": {
                                  "library": "H5P.AdvancedText 1.1",
                                  "params": {
                                    "text": nhap4.cpmverttext_h5p(z.text)
                                    },
                                  "subContentId": str(nhap4.uuid.uuid4()),
                                  "metadata": {
                                    "contentType": "Text",
                                    "license": "U",
                                    "title": "Untitled Text",
                                    "authors": [],
                                    "changes": []
                                  }
                                },
                                "alwaysDisplayComments": False,
                                "backgroundOpacity": 0,
                                "displayAsButton": False,
                                "buttonSize": "big",
                                "goToSlideType": "specified",
                                "invisible": False,
                                "solution": ""
                              } #% (z.x,z.y,z.cx,z.cy,nhap4.cpmverttext_h5p(z.text),str(nhap4.uuid.uuid4()))
                listobj_textfield.append(h5ptext)
                # uidramdom=str(nhap4.uuid.uuid4())
                # h5p_textt={"x":z.x,"y":z.y,"width":z.cx,"height":z.cy,"action":{"library":"H5P.AdvancedText 1.1","params":{"text":,"subContentId":uidramdom,"metadata":{"contentType":"Text","license":"U","title":"Untitled Text","authors":[],"changes":[]}},"alwaysDisplayComments":False,"backgroundOpacity":0,"displayAsButton":False,"buttonSize":"big","goToSlideType":"specified","invisible":False,"solution":""} #% (z.x,z.y,z.cx,z.cy,,str(uuid.uuid4()))
                #print('h5p_text_element',j,h5p_textt)


            listobj_text_slide.append(listobj_textfield)
            listobj_textfield = []
            print('#-------------END di thang obj-------------------')
                #print(j,'z.text',z.text)
    for i,x in enumerate(listobj_text_slide):
        if not x:
            print('slide ',i,'None')
        else:
            print('slide',i)
            for j,data in enumerate(x):
                print(j,'obj_text_h5p',data)
    return listobj_text_slide#,list_none



title='10slide'
folder=r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep'
x=traveobj_slide(title,folder)
with open('axz.json', encoding='utf-8') as fh:
    data = json.load(fh)
a = [i for i, x in enumerate(x) if not x]
print(a)
listobj_text_slidee=dithangobject_ver001(x,arr_none=a)
print('------------#checkkkkkkk--------------------------')
for i, x in enumerate(listobj_text_slidee):
    if not x:
        print('obj ', i, 'None')
    else:
        print('element obj', i)
        for j, data in enumerate(x):
            print(j, 'obj_text_h5p', data)

