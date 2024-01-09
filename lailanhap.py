def  pixelsToEmus( widthPx, heightPx,  resDpiX, resDpiY,  zoomX,  zoomY):

    emusPerInch = 914400;
    emusPerCm = 360000;
    maxWidthCm = 16.51;
    widthEmus = (int)(widthPx / resDpiX * emusPerInch) * zoomX / 100;
    heightEmus = (int)(heightPx / resDpiY * emusPerInch) * zoomY / 100;
    return heightEmus,widthEmus
#print(pixelsToEmus(  8650579 ,1190400,96,96,10,100))
from tempfile import mkstemp
from shutil import move
import os #import fdopen, remove
def replace(file_path, pattern, subst):
    #Create temp file
    fh, abs_path = mkstemp()
    with os.fdopen(fh,'w') as new_file:

        with open(file_path, 'r', encoding='utf-8') as old_file:
            for line in old_file:
                new_file.write(line.replace(pattern, subst))
    #Remove original file
    os.remove(file_path)
    #Move new file
    move(abs_path, file_path)
import shutil
import os
original = r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep_copy\newdataaaaa.json '
target = r'C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p_1_2_deep_copy\newdata_cpo.txt '

#shutil.copyfile (original, target)
# with open(original, 'r', encoding='utf-8') as file:
#     filedata = file.read()
#
#     # Replace the target string
#
# filedata = filedata.replace(r"""\\\%s"""%(''), r"""\%s"""%(''))
# filedata = filedata.replace(r"""\\%s"""%(''), r"""\%s"""%(''))
#
# with open(original, 'w', encoding='utf-8') as file:
#     file.write(filedata)
# Write the file out again
import edit_colon2_
edit_colon2_.replaceslash(original)
#os.rename(target, original)