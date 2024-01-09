import os

# Absolute path of a file
import string

#error vì chưa xử lý được extra- thay thế tạm thủ công... giải nén tay
def edit_name_pptx2rar(title, folder):
    folder = folder + r"""\%s""" % ("")
    old_name = str(folder) + str(title) + ".pptx"
    #new_name = str(folder) + str(title) + "1.rar"

    new_name = str(folder) + str(title) +".pptx"#+ "1.rar"
    new_title = str(title) + "1.rar"
    title_of_folder_exrar = str(title) + "1"

    # Renaming the file
    os.rename(old_name, new_name)
    return new_name, folder, new_title, title_of_folder_exrar
def edit_name_pptx2rar_old(title, folder):
    folder = folder + r"""\%s""" % ("")
    old_name = str(folder) + str(title) + ".pptx"
    new_name = str(folder) + str(title) + "1.rar"
    new_title = str(title) + "1.rar"
    title_of_folder_exrar = str(title) + "1"

    # Renaming the file
    os.rename(old_name, new_name)
    return new_name, folder, new_title, title_of_folder_exrar

#link= r"""C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p-1.2_dêp\""""


