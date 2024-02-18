import aspose.words as aw
import re

# read all the docx files in the current directory
# and convert them to markdown with the same main name and .md extension by aspose.words
# in the same directory
import os

def remove_anchor_tags(filename):
    # 读取文件内容
    with open(filename, 'r') as file:
        content = file.read()

    # 使用正则表达式删除匹配的内容
    cleaned_content = re.sub(r'<a name=".*?"></a>', '', content)

    # 将处理后的内容写回文件
    with open(filename, 'w') as file:
        file.write(cleaned_content)

for file in os.listdir("."):
    if file.endswith(".docx"):
        doc = aw.Document(file)
        file_name, file_extension = os.path.splitext(file)
        doc.save(file_name + ".md", aw.SaveFormat.MARKDOWN)
        remove_anchor_tags(file_name + ".md")





