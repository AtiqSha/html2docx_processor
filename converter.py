from htmldocx import HtmlToDocx
import re
import codecs
import os
import argparse
import glob
import shutil
import sys
from datetime import datetime

start_time = datetime.now()
parser = argparse.ArgumentParser(description='A test program.')

parser.add_argument("--input", help="Folder name that contains all html file to convert")
parser.add_argument("--output", help="Folder name to store all converted html files as docx")

args = parser.parse_args()

input_path = args.input + '\\'
output_path = args.output + '\\'
temp = f'{output_path}/temp/'

if not os.path.exists(temp):
    os.makedirs(temp)
    
if not os.path.exists(output_path):
    os.makedirs(output_path)

files = glob.glob(os.path.join(input_path, '*.html'))
print(f'found {len(files)} html files in {input_path}')
failed = []
for file in files:
    try:
        # print(file)
        html_as_text = codecs.open(file, "r", "utf-8").read()

        #rename file as title/header in html
        titles = re.findall(r"<h1.+?>.+?</h1>", html_as_text)
        sub_search = re.search(r"<h1.+?>(.+?)</h1>",titles[0])
        rename = sub_search.group(1)

        #remove header/title in html
        titles = re.findall(r"<title>.+?</title>", html_as_text)
        splitter = html_as_text.split(titles[0])
        html_as_text = f'{splitter[0]}<!--{titles[0]}-->{splitter[1]}'
        header = re.findall(r"<h1.+?>.+?</h1>", html_as_text)
        splitter = html_as_text.split(header[0])
        html_as_text = f'{splitter[0]}<!--{header[0]}-->{splitter[1]}'
        # print(html_as_text)

        #rearrange list (if needed)

        #rearrange parent folder text
        parent_folder = re.findall(r".+?parentlink.+\n", html_as_text)
        splitter = html_as_text.split(parent_folder[0])
        post_parent_title = re.findall(r"<a class.+?>", parent_folder[0])
        post_proc = parent_folder[0]
        try:
            post_proc = re.sub(post_parent_title[0], '', parent_folder[0])
        except:
            pass
        html_as_text = splitter[0] + '<p></p>\n' + post_proc + splitter[1]

        


        temp_output = f"{temp}temp.html"
        temp_file = open(temp_output, "w")
        temp_file.write(html_as_text)
        temp_file.close()

        new_parser = HtmlToDocx()
        output = f'{output_path}{rename}'
        new_parser.parse_html_file(temp_output, output)
    except:
        failed.append(file)

    # break

try:
    shutil.rmtree(temp)
except OSError as e:
    print("Error: %s - %s." % (e.filename, e.strerror))

end_time = datetime.now()
elapsed = end_time - start_time
print(f'completed [elapsed time: {elapsed}]')
if len(failed) > 0:
    print(f'File failed to be converted : {failed}')
    print('Error: please close the existing .docx files before running the python script')


