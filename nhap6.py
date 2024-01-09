from xml.etree.ElementTree import ElementTree
import zipfile
import edit_colon2_
import os
srcfile = r'C:\Users\Admin\anaconda3\envs\backuptamthoi3\pptx2h5p_1_2_deep_copy_addimages\10slide_cp.rar'
dstfile = r'C:\Users\Admin\anaconda3\envs\backuptamthoi3\pptx2h5p_1_2_deep_copy_addimages\10slide_cp_outtttttttttt.rar'

inzip = zipfile.ZipFile(srcfile, 'r')
outzip = zipfile.ZipFile(dstfile, "w")

# Iterate the input files
for inzipinfo in inzip.namelist():# infolist():
    if inzipinfo.endswith('/'): continue

    #print(inzipinfo)

    # Read input file
    infile = inzip.open(inzipinfo)
    if 'ppt/slides/slide' in inzipinfo:
#         f = zf.open(name)
        print(inzipinfo)
    # if inzipinfo.filename == "test.txt":
    #
        content = infile.read()
    #
        #new_content = str(content, 'utf-8').replace(":", ":")
        #print(content)
    #
        new_content=str("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>""")
    #     # Write content
        outzip.writestr(inzipinfo, new_content)
        tree = ElementTree()

        tree.fromstring(infile.read())
        foos = tree.findall('p_cSld')
        for foo in foos:
            bars = foo.findall('p_spTree')
            print(bars)
        #   for bar in bars:
        #     foo.remove(bar)

        #tree.write(inzipinfo)
    #
    else:  # Other file, dont want to modify => just copy it

        content = infile.read()
        outzip.writestr(inzipinfo, content)

inzip.close()
outzip.close()
#os.remove(dstfile)
# tree = ElementTree()
#
# tree.parse(f.read())
# foos = tree.findall('<p_cSld')
# for foo in foos:
#   bars = foo.findall('p_spTree')
#   for bar in bars:
#     foo.remove(bar)
#
# tree.write(name)