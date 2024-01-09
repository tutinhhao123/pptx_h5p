import os, sys
def edit_name_xml2txt(title, folder):
    link_rel_old =folder + r"""_copy\ppt\slides\_rels\%s""" % ("")+str(title) + ".xml.rels"
    link_rel_new = folder + r"""_copy\ppt\slides\_rels\%s""" % ("") + str(title) + ".txt"

    folder = folder + r"""_copy\ppt\slides\%s""" % ("")
    old_name = str(folder) + str(title) + ".xml"
    #new_name = str(folder) + str(title) + "1.rar"

    new_name = str(folder) + str(title) +".txt"#+ "1.rar"
    new_title = str(title) + "1.rar"
    title_of_folder_exrar = str(title) + "1"

    # Renaming the file
    os.rename(link_rel_old, link_rel_new)
    os.rename(old_name, new_name)
    return new_name, folder, new_title, title_of_folder_exrar,link_rel_new
def edit_name_txt2xml(title, folder):
    link_rel_new = folder + r"""_copy\ppt\slides\_rels\%s""" % ("") + str(title) + ".xml.rels"
    link_rel_old = folder + r"""_copy\ppt\slides\_rels\%s""" % ("") + str(title) + ".txt"

    folder = folder + r"""_copy\ppt\slides\%s""" % ("")
    old_name = str(folder) + str(title) + ".txt"
    #new_name = str(folder) + str(title) + "1.rar"

    new_name = str(folder) + str(title) +".xml"#+ "1.rar"
    new_title = str(title) + "1.rar"
    title_of_folder_exrar = str(title) + "1"

    # Renaming the file
    os.rename(old_name, new_name)
    # os.rename(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1_copy\ppt\slides\slide5.txt', new_name)
    os.rename(link_rel_old, link_rel_new)
    return new_name, folder, new_title, title_of_folder_exrar,link_rel_new
def edittext_(linkfile):
    with open(linkfile, 'r',encoding='utf-8') as file:
        filedata = file.read()

    # Replace the target string
    filedata = filedata.replace(':', '_')

    # Write the file out again
    with open(linkfile, 'w',encoding='utf-8') as file:
        file.write(filedata)
def editname2txt(link):
    print("Thu muc gom: %s" % os.listdir(os.getcwd()))

    # thay ten thu muc ''vietjackdir"
    os.rename("vietjackdirectory", "slide5.xml")

    print("Thay ten thanh cong.")

    # Liet ke cac thu muc sau khi thay ten "vietjackdir"
    print("Thu muc gom: %s" % os.listdir(os.getcwd()))

# Liet ke cac thu muc
if __name__ == "__main__":
    print( "Thu muc gom: %s"%os.listdir(os.getcwd()))

    # thay ten thu muc ''vietjackdir"
    a,b,c,d,e=edit_name_xml2txt('slide5',r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1')
    # Read in the file
    # print('',e)
    print(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1_copy\ppt\slides\_rels')
    edittext_(a)
    edittext_(e)
    # os.rename(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1_copy\ppt\slides\_rels\slide5.txt',
    #           r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1_copy\ppt\slides\_rels\slide5.xml.rels')

    # os.rename(r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1_copy\ppt\slides\_rels\slide5.xml.rels',r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1_copy\ppt\slides\_rels\slide5.txt')
    edit_name_txt2xml('slide5',r'C:\Users\Admin\.conda\envs\backuptamthoi2\pptx2h5p_1_2_deep\a1')