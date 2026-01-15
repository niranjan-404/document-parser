from docling.document_converter import DocumentConverter
from langchain.text_splitter import MarkdownHeaderTextSplitter 
import fitz 
import docx
from io import BytesIO
from PIL import Image
import os
from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser, JsonOutputParser
from langchain_community.document_loaders.word_document import Docx2txtLoader
from langchain_community.document_loaders.text import TextLoader
from langchain_community.document_loaders.pdf import PyPDFLoader
from langchain_community.document_loaders.powerpoint import UnstructuredPowerPointLoader
from langchain_core.output_parsers import BaseOutputParser
from langchain.tools.base import StructuredTool
from langchain_core.output_parsers import BaseOutputParser
from typing import TypedDict, Annotated, List, Union
from langchain_core.agents import AgentAction, AgentFinish
from langchain_core.messages import BaseMessage
from langchain_core.messages import AIMessage, HumanMessage
import base64
from mimetypes import guess_type
from langchain_core.prompts import (
    ChatPromptTemplate,
    MessagesPlaceholder,
    PromptTemplate,)
from langchain.schema.document import Document
from datetime import timezone,date
from datetime import datetime
from pathlib import Path
from typing import List
from langchain.schema import Document
import fitz  
from docling.document_converter import DocumentConverter
from langchain.text_splitter import MarkdownHeaderTextSplitter  
import tempfile
import os
from pathlib import Path
import operator

class DocumentParser():
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.converter = DocumentConverter()

    def load_document(self) -> List[Document]:
        if self.file_path.endswith('.pdf'):
            return self.load_pdf_document()
        elif self.file_path.endswith('.docx'):
            return self.load_docx_document()
        elif self.file_path.endswith('.txt'):
            return self.load_text_document()
        elif self.file_path.endswith('.pptx'):
            return self.load_pptx_document()
        elif self.file_path.endswith('.md'):
            return self.load_markdown_document()
        else:
            raise ValueError("Unsupported file type")

    def load_pdf_document(self):
        loader = PyPDFLoader(self.file_path)
        return loader.load()

    def load_docx_document(self):
        loader = Docx2txtLoader(self.file_path)
        return loader.load()

    def load_text_document(self):
        loader = TextLoader(self.file_path)
        return loader.load()

    def load_pptx_document(self):
        loader = UnstructuredPowerPointLoader(self.file_path)
        return loader.load()
    
    def find_image_indices_with_details(self,text, image_marker="<!-- image -->"):
        """
        Enhanced version that provides more details about each occurrence.
        
        Args:
            text (str): The text to search in
            image_marker (str): The marker to search for (default: "<!-- image -->")
        
        Returns:
            list: List of dictionaries with detailed information about each occurrence
        """
        details = []
        start = 0
        marker_length = len(image_marker)
        occurrence_count = 0
        
        while True:
            index = text.find(image_marker, start)
            
            if index == -1:
                break
            
            occurrence_count += 1
            end_index = index + marker_length

            line_number = text[:index].count('\n') + 1
            
            last_newline = text.rfind('\n', 0, index)

            column = index - last_newline - 1 if last_newline != -1 else index
            
            details.append({
                'occurrence': occurrence_count,
                'start_index': index,
                'end_index': end_index,
                'line_number': line_number,
                'column': column
            })
            
            start = end_index
        
        return details

    def load_markdown_document(self,file_path: str = None) -> List[Document]:

        if file_path is None:
            return []
        
        self.file_path = file_path

        doc = fitz.open(self.file_path)
        
        file_name = Path(self.file_path).name

        parsed_docs = []

        first_twelve_pages = ""

        for page_number, page in enumerate(doc, start=1):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                        temp_page_path = tmp_file.name

                page_content = {}
                page_document = fitz.open()  
                page_document.insert_pdf(doc, from_page=page_number - 1, to_page=page_number - 1)
                page_document.save(temp_page_path)
                page_document.close()

                conversion_result = self.converter.convert(temp_page_path)

                page_markdown = conversion_result.document.export_to_markdown()

                if os.path.exists(temp_page_path):
                    os.remove(temp_page_path)

                page_content["text"] = page_markdown

                page_content["page_number"] = page_number

                page_content["source"] = self.file_path

                image_details = self.find_image_indices_with_details(page_markdown)

                page_content["images"] = self.extract_images_from_file(page_number,self.file_path,image_details,page_content)
                
                parsed_docs.append(page_content)

        return parsed_docs 
    
    def extract_images_from_file(self, page_number,image_details:list,page_content:dict):


            with tempfile.TemporaryDirectory() as temp_dir:
                output_dir = os.path.join(temp_dir, "extracted_images")
                os.makedirs(output_dir, exist_ok=True)
            
            extracted_images = []

            extracted_text = page_content["text"]

            number_of_images = len(image_details)

            if self.file_path.lower().endswith(".pdf"):
                doc = fitz.open(self.file_path)
                if page_number < 1 or page_number > len(doc):
                    raise ValueError("Invalid page number for PDF.")
                
                page = doc[page_number - 1] 
                images = page.get_images(full=True)

                for i, img in enumerate(images):
                    xref = img[0]

                    base_image = doc.extract_image(xref)

                    image_bytes = base_image["image"]

                    image_ext = base_image["ext"]
                    
                    image_path = os.path.join(output_dir, f"pdf_page{page_number}_img{i}.{image_ext}")

                    image_details[i]["imagebase64"] = base_image["image"]
                    
                    image_details[i]["image_type"] = base_image["ext"]

                    img_name =f"Image_{str(page_number)}_{str(i)}.{base_image['ext']}"

                    image_details[i]["image_name"] = img_name

                    page_content["text"] = page_content["text"][image_details[i]["start_index"]:image_details[i]["end_index"]] = img_name
                
                page_content["images"] = image_details

            elif self.file_path.lower().endswith(".docx"):
                doc = docx.Document(self.file_path)
                images = doc.inline_shapes

                if page_number != 1:
                    raise ValueError("DOCX files do not have pages like PDFs. Extracting from the whole document.")

                for i, shape in enumerate(images):
                    if shape._inline.graphic.graphicData.pic.blipFill:
                        blip = shape._inline.graphic.graphicData.pic.blipFill.blip
                        rID = blip.embed
                        image_part = doc.part.related_parts[rID]
                        image_bytes = image_part.blob

                        image = Image.open(BytesIO(image_bytes))
                        image_ext = image.format.lower()  

                        image_path = os.path.join(output_dir, f"docx_img{i}.{image_ext}")
                        image.save(image_path)
                        extracted_images.append(image_path)

            else:
                raise ValueError("Unsupported file format. Please provide a PDF or DOCX file.")

            return extracted_images
