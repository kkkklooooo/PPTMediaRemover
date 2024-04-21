import zipfile,tempfile
import numpy as np
import re,os,argparse
from colorama import init,Fore,Back,Style
import shutil

parser=argparse.ArgumentParser(description="PPT Media Remover. Reduce the size of ppt file")
parser.add_argument("-p","--path",help="path of the file or folder to be processed",type=str,required=True)
parser.add_argument("-k","--key",help="keywords to be removed",type=str,nargs="*",default=".mp4 .wmv")
parser.add_argument("-c","--compress",help="compress the media rather than delete",action="store_true")
args=parser.parse_args()
if os.path.isfile(args.path):
    print(Fore.GREEN+Style.BRIGHT+f"Working mode: Single file\n\t{args.path}")
    print(Style.RESET_ALL)
elif os.path.isdir(args.path):
        print(Fore.GREEN+Style.BRIGHT+f"Wroking mode: Folder\n\t{args.path}")
        print(Style.RESET_ALL)
else:
    print(Fore.RED+Style.BRIGHT+"Invalid path")
    print(Style.RESET_ALL)

print(Fore.GREEN+Style.BRIGHT+f"\nKeywords: {args.key}")
print(Style.RESET_ALL)

def Debug(str):
    print(Fore.RED+Style.BRIGHT+"[DEBUG]"+str)
    print(Style.RESET_ALL)

input=input("Press Enter to continue>>")
#key=['.mp4',".wmv"]
key=args.key.split()
key_f=r""
for i in key:
    key_f+=(f"{i}|")
key_f=re.compile(key_f[:-1])
Debug(f"{key}")
path=args.path


def MultiProcess(path):
    count=0
    for file in os.listdir(path):
        count+=1
        print(Fore.YELLOW+Style.BRIGHT+f"[Info] Processing {count}/{len(os.listdir(path))}")
        tpf=tempfile.mkdtemp()
        file_full=os.path.join(path,file)
    #path=r"E:\临时分类文件\2023-03\生活中的圆周运动20230224.pptx"
        O_dir,o_name=os.path.split(file_full)
          
        Process(file_full,O_dir,o_name,tpf)
        

def Process(path,O_dir,o_name,tpf):
    try :
        f=zipfile.ZipFile(path,"r")
    except:
        print(Fore.YELLOW+Style.BRIGHT+f"[Warning] {o_name} is not a PPT file,Skip...")
        return False
    b=[ key_f.search(file.filename)==None for file in f.filelist]
    
    data=np.array(f.filelist)
    f.extractall(tpf,members=data[b].tolist())
    f.close()
    n_data=[]
    for dir,_,flist in os.walk(tpf):
        for file in flist:
            n_data.append(os.path.join(dir,file))
    with zipfile.ZipFile(f"{O_dir}\\Processed_{o_name}","x") as f_new:
        for file in n_data:
            f_new.write(f"{file}",f"{os.path.relpath(file,tpf)}")
    f_new.close()
    print(Fore.GREEN+Style.BRIGHT+f"[Info] {o_name} processed")
    print(Style.RESET_ALL)
    os.remove(path)
    
    shutil.rmtree(tpf)
    print(Fore.GREEN+Style.BRIGHT+f"[Info] {tpf} removed")
    print(Style.RESET_ALL)
    return True
if os.path.isfile(path):
    tpf=tempfile.mkdtemp()
    Process(path,os.path.split(path)[0],os.path.split(path)[1],tpf)
if os.path.isdir(path):
    MultiProcess(path)