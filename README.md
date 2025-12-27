# PDF to MS Word Replication using Python

## Overview
This project recreates a given legal PDF form into an MS Word document using Python.

## Libraries Used
- python-docx
- lxml (via OXML)
- Python 3.x

## Approach
1. Analyzed the PDF layout including headings, spacing, and table structure
2. Recreated the document from scratch using python-docx
3. Used Word tables with merged cells to match the PDF layout
4. Applied custom paragraph spacing using OXML for precise formatting
5. Preserved template placeholders and conditional blocks

## How to Run
```bash
pip install -r requirements.txt
python main.py
