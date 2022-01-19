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


st.title('Wordコメント著者情報削除')
st.markdown(f"Wordファイル（docx）をアップロートしてください。<br>処理終了後、ダウンロードリンクが表示されます。", unsafe_allow_html=True)
uploaded_file = st.file_uploader('下欄にドラッグ＆ドロップできます。', type='docx')


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
                        comments_new += re.sub(r'w:author="[^"]*"', "w:author=\"Anonymous Author\"", comments.decode())
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
            href = f'<a href="data:application/docx;base64,{encoded_string}" download="{download_filename}">download</a>'
            st.markdown(f"ダウンロードする {href}", unsafe_allow_html=True)

        try:
            os.remove(output_filename)
        except:
            pass

    except Exception as e:
        st.markdown(f"<b>エラーが発生しました: {str(e)}</b>", unsafe_allow_html=True)

