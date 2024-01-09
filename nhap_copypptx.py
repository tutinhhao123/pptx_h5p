import shutil
import os
src=r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep_copy\10slide.pptx'
dst=r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep_copy\10slide_copy1.pptx'
shutil.copyfile(src, dst)
#os.remove(dst)


pres = Presentation(dst) # name of your pptx
slide = pres.slides[0] # or whatever slide your pictures on
pic  = slide.shapes[0] # select the shape holding your picture and you want to remove

pic = pic._element
pic.getparent().remove(pic) # delete the shape with the picture

pres.save("test.pptx") # save changes




# 2nd option
#shutil.copy(src, dst)