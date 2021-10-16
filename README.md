### html2docx_processor (intended specific usage)

To convert html file into docx file.

####Install external package:
```
pip install htmldocx
```

Package retrieved from:https://pypi.org/project/htmldocx/
Refer full package code on GitHub : https://github.com/pqzx/html2docx.git

####Preprocessing before converting to docx:
1) Rename file name as in header
2) Remove/comment out header and title to avoid displaying it in docx
3) Remove href and adding new paragraph for better text arrangement


####Run script:
```
cd <path/to/converter.py>
python converter.py --input "<path/to/input/folder>" --output "<path/to/output/folder>"
```
Double quote (" ") is needed for path in case that the path consists of spaces.
