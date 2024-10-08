import os
import re
import logging

from PyPDF2 import PdfMerger

LOGGER = logging.getLogger()

def main():
    configure_logger(LOGGER, 'debug')
    mergelist = get_merge_list()
    combined_pdf = "Clyde-AA Reciept.pdf"  # Output combined PDF file
    combine_pdfs(mergelist, combined_pdf)
    LOGGER.info("PDFs combined and inserted into PowerPoint presentation successfully.")

def configure_logger(logger, level = 'info'):
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler = logging.StreamHandler()
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    log_level = level.upper()
    logger.setLevel(log_level)    

def get_merge_list():
    folder = os.path.join(os.getcwd(),'files_to_merge')
    mergelist = []
    for pdf in os.listdir(folder):
        if re.search('.pdf$', pdf):
            mergelist.append(os.path.join(folder, pdf))
    LOGGER.debug('Merge List Created')
    LOGGER.debug(mergelist)
    return mergelist

def combine_pdfs(pdf_files, output_file):
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    merger.write(output_file)
    merger.close()

if __name__ == "__main__":
    main()