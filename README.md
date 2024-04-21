# PPT Media Remover
## Overview
PPT Media Remover 是一个 Python 脚本，旨在通过移除或压缩 PowerPoint (.pptx) 文件中的特定媒体文件（如 .mp4 和 .wmv），以减小文件大小。该脚本支持处理单个文件或整个目录中的 PPT 文件。

## Features
媒体文件筛选：根据指定的关键字列表（默认为 .mp4 和 .wmv）识别并处理目标媒体文件。
删除或压缩：可以选择删除指定媒体文件以减小 PPT 文件大小，或者选择压缩媒体文件（默认行为）以保留原始内容但节省空间。
批量处理：能够递归遍历指定目录，处理其中的所有 PPT 文件。
清晰的命令行输出：提供丰富的颜色和样式输出，包括处理模式、文件路径、进度信息、警告与错误提示等。
Requirements
Python 3.x
下列第三方库（已包含在 requirements.txt 中）：
numpy
colorama
Installation
在项目根目录下，通过以下命令安装所需依赖：

```bash
pip install -r requirements.txt
Usage
Command Line Arguments
```
运行 python ppt_media_remover.py 并传入以下参数：
```
usage: ppt_media_remover.py [-h] [-p PATH] [-k KEY [KEY ...]] [-c]

PPT Media Remover. Reduce the size of ppt file

optional arguments:
  -h, --help            show this help message and exit
  -p PATH, --path PATH  path of the file or folder to be processed
  -k KEY [KEY ...], --key KEY [KEY ...]
                        keywords to be removed
  -c, --compress        compress the media rather than delete
```
参数详解：
-p / --path 必填：指定要处理的 PPT 文件或包含 PPT 文件的目录路径。
-k / --key 可选：指定要移除或压缩的媒体文件扩展名关键词，多个关键词之间以空格分隔。默认为 .mp4 .wmv。
-c / --compress 可选：启用此选项时，将压缩而非删除指定媒体文件。默认行为为删除。
示例
处理单个 PPT 文件
```bash
python ppt_media_remover.py -p path/to/some_presentation.pptx
```
处理目录中的所有 PPT 文件
```bash
python ppt_media_remover.py -p path/to/presentations_folder -c
```
在这个例子中，脚本将遍历 presentations_folder 目录及其子目录，对所有 PPT 文件进行处理，并且压缩而非删除指定媒体文件。
