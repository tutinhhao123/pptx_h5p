import json
from nhap4 import traveobj_slide
from nhap4 import convert_donvi_text_xycxcy_2_h5p
from nhap4 import converttext_h5p_mini
from nhap4 import cpmverttext_h5p
from nhap4 import replacestring
import uuid
def print_op(content):

    for i,a in enumerate(content["presentation"]["slides"]) :
        print('slide',i)
        #print( a)
        print(count_keys(a['elements']))
        for j,elements in enumerate(a['elements']):

            print(elements)
    return content
def count_keys(dict):
    count = 0
    for i in enumerate(dict):
        count += 1
    return count
list01 = [1, 2, 3]
list02 = ['x','y']
for i, (a, b) in enumerate(zip(list01, list02)):
    print (i, a, b)
def delete_text_element1_inslide(content):
    for i, a in enumerate(content["presentation"]["slides"]):
        print('slide', i)
        # print( a)
        print(count_keys(a['elements']))
        for j, elements in enumerate(a['elements']):
            # if(j>0):
            print(j)
            if j > 0 and j < 2:
                del a['elements'][j]
    return content
def add_rong_elements_0(list):
    for i, dataa in enumerate(list):
        if not i:
            print("k ton tai")
        else:
            dataa.insert(0, '')
            print('slide i', i)
            for numb, j in enumerate(dataa):
                # j.insert(0,'')
                if not j:
                    print('obj', numb, ' is none')
                else:
                    print(numb, j)

    return list
def edit_elements_slide(data,x):
    a = [i for i, x in enumerate(x) if not x]
    print(a)
    for i, (slide, arr_obj) in enumerate(zip(data["presentation"]["slides"], x)):
        print('#------------------------start with h5p add text')
        if (True == True):
            if i not in a:
                print('slide', i)
                # print("gia tri cua arrayobj",i,arr_obj)
                # print('#-------------------------text obj.x.y-------------------')

                # get element to add param text to json
                # edit element slides ♥ god bles u
                # tính lại x,y, cx,cy
                # viet method
                # convert_donvi_text_xycxcy_2_h5p()
                # print('#-------------------------end text obj.x.y-------------------')
                # for j,data in enumerate(arr_obj):

                for j, (elements, data) in enumerate(zip(slide['elements'], arr_obj)):
                    # add element text cho tung slide

                    # tính lại x,y, cx,cy
                    #         viet method
                    if (True):
                        if data == '':
                            print('data', j, ' is none\n')
                        else:

                            # ----------------------------------------- phai xu ly duoc list index
                            print('#-----------start convert x y cx cy---------------')
                            print('gia tri array[', i, '],[', j, ']')

                            print(data.x, data.y,data.cx,data.cy,data.text)
                            data.x, data.y, data.cx, data.cy = convert_donvi_text_xycxcy_2_h5p(data.x, data.y, data.cx,
                                                                                               data.cy)
                            print(data.x, data.y,data.cx,data.cy,data.text)
                            print('slide[', i, ']', 'elements[', j, ']', elements)

                            if(False):
                                print('gia tri cua elements object')
                                print(elements['x'], # ['x']
                                elements['y'] , # ['y']
                                elements['width'] ,  # ['cx']
                                elements['height']  # ['cy'])
                                      ,
                                      elements["action"]["params"]
                                      )
                                params = elements["action"]["params"]
                                params.update({"text": cpmverttext_h5p(data.text)})
                                print('new value parm text',params)
                                print('value elements["action"]["subContentId"] ',elements["action"]["subContentId"] )
                                elements["action"]["subContentId"] = str(uuid.uuid4())
                                print('value elements["action"]["subContentId"] ', elements["action"]["subContentId"])
                            print('#-----------end convert x y cx cy-----------------')
                            # params = elements["action"]["params"]
                            # params.update({"text": replacestring(cpmverttext_h5p(data.text))})
                            # print('text',params)
                            # slides[0]["elements"][1]
                            # elements = slides[i]["elements"][j+1]
                            # checkslide_arr(elements,data)
                            if (False ):
                                elements['x'] = data.x  # ['x']
                                elements['y'] = data.y  # ['y']
                                elements['width'] = data.cx  # ['cx']
                                elements['height'] = data.cy  # ['cy']

                                params = elements["action"]["params"]
                                # add new image filename
                                # ---------------------------
                                #params["text"] = replacestring(cpmverttext_h5p(cpmverttext_h5p(data.text)))
                                params.update({"text": cpmverttext_h5p(data.text)})
                                # ---------------------------
                                # add random uuid
                                elements["action"]["subContentId"] = str(uuid.uuid4())
    return data
# f = open('axz.json','r')
# contents = f.read()
# print(contents)
#f.close()
title='10slide'
folder=r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep'
x=traveobj_slide(title,folder)
with open('axz.json', encoding='utf-8') as fh:
    data = json.load(fh)

#print(data)
print(count_keys(data))
numb=0
# print_op(data)
        #print(elements)
    #print(data.count(i['elements']))
#data = print_op(content=delete_text_element1_inslide(data))

#data = delete_text_element1_inslide(data)
print(type(data))
print('new voalues')
#print_op(data)
print(type(data))
print(type(x))
#traveobj_slide()
# thislist = ["apple", "banana", "cherry"]
# #thislist.insert(0, "")
# print(thislist)
# for i in x:
#     if not i:
#         print("k ton tai")
#     else:
#         #i.insert(0,'')
#         print('value i',i)
#         for j in i:
#             #j.insert(0,'')
#             print(j)
print(type(x))

#print(x)
add_rong_elements_0(x)

#data = delete_text_element1_inslide(data)
#n=[1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20]

print('new values of dataa')
# out_file = open("newdata.txt", "w")
#
# json.dump(data, out_file, indent=6)
#
# out_file.close()
# f = open("newdata.txt", "w")
# f.write(json.dumps(data, f, indent=6))
# f.close()
print(type(data))
# for i in data["presentation"]["slides"]:
#     for j in i:
#         print (j)