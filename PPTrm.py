import zipfile,tempfile
import numpy as np
import re,os,argparse
from colorama import init,Fore,Back,Style
import shutil,subprocess

parser=argparse.ArgumentParser(description="PPT Media Remover. Reduce the size of ppt file")
parser.add_argument("-p","--path",help="path of the file or folder to be processed",type=str,required=True)
parser.add_argument("-k","--key",help="keywords to be removed",type=str,nargs="*",default=".mp4 .wmv")
parser.add_argument("-c","--compress",help="compress the media rather than delete",action="store_true")
args=parser.parse_args()
compress=args.compress
if compress:
    ffmpegV=os.popen("ffmpeg -version").read().split("\n")[0]
    print(Fore.GREEN+Style.BRIGHT+f"Using:{ffmpegV}")
    print(Style.RESET_ALL)
    if ffmpegV.find("ffmpeg version")==-1:
        print(Fore.RED+Style.BRIGHT+"ffmpeg not found, please install ffmpeg or disable compress mode")
        print(Style.RESET_ALL)
        exit()
if os.path.isfile(args.path):
    print(Fore.GREEN+Style.BRIGHT+f"Working mode: Single file\n\t{args.path}")
    print(Style.RESET_ALL)
elif os.path.isdir(args.path):
        print(Fore.GREEN+Style.BRIGHT+f"Wroking mode: Folder\n\t{args.path}")
        print(Style.RESET_ALL)
else:
    print(Fore.RED+Style.BRIGHT+"Invalid path")
    print(Style.RESET_ALL)

print(Fore.GREEN+Style.BRIGHT+f"\nKeywords: {args.key}\nCompress: {compress}")
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
    p=[ key_f.search(file.filename)!=None for file in f.filelist]
    data=np.array(f.filelist)
    if not compress:
        f.extractall(tpf,members=data[b].tolist())
    else:
        f.extractall(tpf,members=data.tolist())
        for i,file in enumerate(data[p].tolist()):
            print(f"[Info] Compressing {i+1}/{len(data[p])}")
            Compress(os.path.join(tpf,file.filename))
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
    try:
        shutil.rmtree(tpf)
        os.remove(path)
    #----------
        
        print(Fore.GREEN+Style.BRIGHT+f"[Info] {tpf} removed")
        print(Style.RESET_ALL)
    except:
        print(Fore.RED+Style.BRIGHT+f"[Error] {tpf}or{path} not removed")
    return True

def Compress(i_path):
    video_extensions = ['.mp4', '.avi', '.mkv', '.mov', '.wmv']
    
    # 检查文件扩展名是否为视频文件
    if not any(i_path.lower().endswith(ext) for ext in video_extensions):
        print(f"{i_path} 不是视频文件。")
        return None
    
    # 获取文件名和目录
    file_name = os.path.basename(i_path)
    base_name, ext = os.path.splitext(file_name)
    output_path = os.path.join(os.path.dirname(i_path), f"{base_name}_out{ext}")
    
    # 构建ffmpeg命令
    ffmpeg_cmd = [
        "ffmpeg", "-i", i_path,
        "-c:v", "libx265",  # 使用H.265编码器进行高压缩
        "-crf", "51",       # 压缩率控制，数值越小质量越好但文件越大，通常18-28之间
        "-preset", "fast", # 压缩速度与质量平衡，可选快速如"fast"或高质量如"veryslow"
        output_path
    ]
    
    try:
        # 执行ffmpeg命令
        subprocess.run(ffmpeg_cmd, check=True)
        print(Fore.GREEN+Style.BRIGHT+f"[Info] 视频文件已压缩")
        print(Style.RESET_ALL)
        os.remove(i_path)
        os.rename(output_path, i_path)
    except subprocess.CalledProcessError as e:
        print(Fore.RED+Style.BRIGHT+f"[Error] 压缩视频时发生错误: {e},原视频文件无更改")
        print(Style.RESET_ALL)
        return None
    
    return i_path
if os.path.isfile(path):
    tpf=tempfile.mkdtemp()
    Process(path,os.path.split(path)[0],os.path.split(path)[1],tpf)
if os.path.isdir(path):
    MultiProcess(path)