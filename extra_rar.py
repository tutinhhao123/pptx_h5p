import patoolib
from unrar import rarfile
#patoolib.extract_archive("foo_bar.rar", outdir="path here")
def extrra(namefile,path):
    patoolib.extract_archive(namefile, outdir=path)
if __name__ == "__main__":
    path = r"C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p-1.2_dêp"
    pathfile = r"C:\Users\Admin\anaconda3\envs\backup_1\pptx2h5p-1.2_dêp"
    extrra("a1.rar",path)

