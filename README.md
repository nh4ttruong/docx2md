# Convert .docx to .md (Github)
The python script uses for convert multi .docx file (folder .docx files) to each folder .md files including media source

I wrote it to optimal the time to convert the lab files (.docx) to .md for git up to Github. Note that I just make it for Windows system. If you want to use it on Linux, you have to install as requirement file and change *dir* path!

## Requirements
I put the [requirement file here](./requirements.txt).
- docx==0.2.4 (python-docx)
- pandoc==2.1

Note: If you can run it on Windows, try to access [pandoc](https://pandoc.org/installing.html) and download it!

## Usage
You can specify one or more folder to convert by edit `paths` list still *'DONEEEEE'*:
```python
new_folder = "convert_done" #destination after convert
paths = ['C:\\Users\\truon\\Desktop\\HK6\\NT213-BMW\\BT3.1\\only-docx\\', 'C:\\Users\\truon\\Desktop\\HK6\\NT213-BMW\\BT4.1\\only-docx\\', 'C:\\Users\\truon\\Desktop\\HK6\\NT213-BMW\\BT5.1\\only-docx\\'] #list path to convert multi folder
convert(paths)
```