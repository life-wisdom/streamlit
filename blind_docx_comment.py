import os
import re
import io
import zipfile
import tempfile
import hashlib
import base64
import streamlit as st
from zipfile import ZipFile
from docx import Document


st.title('Wordコメント作者名を削除・変更する (Delete or modify the author name from comments in a docx file.)')

st.markdown(f"Wordファイルに含まれるコメントから、作者名を削除・変更します。 This web service deletes or modifies the author name from comments in a docx file.", unsafe_allow_html=True)
st.markdown(f"下欄に変更後の名前を入力してください（作者名を削除する場合は空欄）。 Input the name to which you want to change the author name below (Leave blank if you want to delete the author name).", unsafe_allow_html=True)

user_input = st.text_input("作者名（author name）", "Anonymous Author")

st.markdown(f"Wordファイル（docx）をアップロートしてください。Upload a Word (docx) flie.<br>処理終了後、ダウンロードリンクが表示されます。Download link will show up after the process.", unsafe_allow_html=True)
uploaded_file = st.file_uploader('下欄にドラッグ＆ドロップできます。(Drop a file below.) ', type='docx')


if uploaded_file is not None:

    try:
        doc = Document(uploaded_file)

        core_properties = doc.core_properties
        meta_fields= ["author", "category", "last_modified_by", "comments", "content_status", "identifier", "keywords", "language", "subject", "title", "version"]
        for meta_field in meta_fields:
            setattr(core_properties, meta_field, "")
        doc.save(uploaded_file.name)

        output_filename = hashlib.sha224(uploaded_file.name.encode()).hexdigest()
        
        # filename
        srcfile = uploaded_file.name  # docx file
        dstfile = output_filename

        with zipfile.ZipFile(srcfile) as inzip, zipfile.ZipFile(dstfile, "w") as outzip:
            # Iterate the input files
            for inzipinfo in inzip.infolist():
                
                # Read input file
                with inzip.open(inzipinfo) as infile:

                    if inzipinfo.filename.startswith("word/comments.xml"):
                        
                        comments = infile.read()
                        comments_new = str()
                        comments_new += re.sub(r'w:author="[^"]*"', f"w:author=\"{user_input}\"", comments.decode())
                        outzip.writestr(inzipinfo.filename, comments_new)

                    else: # Other file, dont want to modify => just copy it

                        outzip.writestr(inzipinfo.filename, infile.read())
        try:
            os.remove(uploaded_file.name)
        except:
            pass

        download_filename = uploaded_file.name
        with open(output_filename, mode="rb") as f:
            content = f.read()
            encoded_string = base64.b64encode(content)
            encoded_string = encoded_string.decode()
            href = f'<a href="data:application/docx;base64,{encoded_string}" download="{download_filename}">{download_filename}</a>'
            st.markdown(f"Download: {href}", unsafe_allow_html=True)

        try:
            os.remove(output_filename)
        except:
            pass

    except Exception as e:
        st.markdown(f"<b>Error: {str(e)}</b>", unsafe_allow_html=True)

