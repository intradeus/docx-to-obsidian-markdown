import os, sys, argparse
import subprocess
import uuid
import re
from PIL import Image, ImageChops
import shutil
from pathlib import Path

OBSIDIAN_ATTACHMENTS = ""
LIBREOFFICE_EXECUTABLE = ""
DEFAULT_MEDIA_OUTPUT = ""
COPY_ALL_FILES = False
FILES_TO_COPY = []
POWERPOINT_TO_PDF = False
LINK_ALL_FILES = False
LINK_TO_UNSUPPORTED_FILES = []


def convert_files(output_dir, root, files):
    """ Single directory batch conversion/importation"""
    for filename in files:
        _, file_extension = os.path.splitext(filename)
        input_file_path = os.path.join(root, filename)
        filename = filename.replace("#", "-").replace("%", "-") # Clean characters that cause issue
        output_file_path = os.path.join(output_dir, filename)
        to_delete = []
        # If file is .doc, convert it to docx
        if os.path.isfile(input_file_path) and file_extension == ".doc":
            input_file_path = convert_doc_to_docx(input_file_path)
            file_extension = ".docx"
            output_file_path = re.sub("\.doc$", ".docx", output_file_path) # Replace the filename from .doc to .docx
            to_delete.append(input_file_path)

        if os.path.isfile(input_file_path) and file_extension == ".docx":
            output_file_path = re.sub("\.docx$", ".md", output_file_path) # Replace the filename from .docx to .md
            Path(output_dir).mkdir(parents=True, exist_ok=True) # Create necessary directories
            print("Copying file : " + input_file_path)
            print("To : " + output_file_path)
            run_conversion(root, input_file_path, output_file_path, file_extension)

        # If file is part of the files-to-copy arguments, just copy it to the obsidian vault
        if os.path.isfile(input_file_path) and (file_extension in FILES_TO_COPY or COPY_ALL_FILES) and file_extension not in [".docx",".doc"]:
            if file_extension in [".pptx", ".ppt"] and POWERPOINT_TO_PDF:
                print("Converting : " + input_file_path)
                print("To : PDF")
                input_file_path = convert_pp_to_pdf(input_file_path)
                output_file_path = re.sub("\.pptx?$", ".pdf", output_file_path)
                to_delete.append(input_file_path)

            elif file_extension in LINK_TO_UNSUPPORTED_FILES or LINK_ALL_FILES:
                print("Creating markdown link file")
                # Create a new markdown file with a link to this file
                new_md_file_path = output_file_path + ".md"
                file_content = "![Link](./" + filename.replace(" ", "%20") + ") \n"
                create_markdown_link_file(new_md_file_path, file_content)

            print("Copying file : " + input_file_path)
            print("To : " + output_file_path)
            Path(output_dir).mkdir(parents=True, exist_ok=True) # Create necessary directories
            shutil.copyfile(input_file_path, output_file_path)

        if len(to_delete) > 0:
            for item in to_delete:
                print("Deleting created artefact : " + item)
                try:
                    os.remove(item)
                except Exception as e:
                    print("Error while deleting: " + e)

        print("---")


def run_conversion(dir, input_file_path, output_file_path, file_extension):
    global DEFAULT_MEDIA_OUTPUT
    DEFAULT_MEDIA_OUTPUT = os.path.join(dir, 'media') # Default pandoc output for images is a folder called /media
    # markdown_mmd
    # Use pandoc to convert the docx to markdown, using standalone tags and extracting media to our dir folder (it will create a media subfolder)
    try:
        cmd_str = "pandoc -f " + file_extension[1:] + " -t markdown_mmd -s --wrap=none --extract-media=\""+ dir + "\" \"" + input_file_path + "\" -o \"" + output_file_path + "\""
        subprocess.run(cmd_str, shell=True)
    except Exception as e:
        print("Error while converting from docx to markdown : " + e)

    clean_file(output_file_path)  # Remove any known artefacts from the file

    clean_images(output_file_path) # Convert images and rename them in each file
    

def clean_images(md_file_path):
    if(os.path.isdir(DEFAULT_MEDIA_OUTPUT)):
        for image in os.listdir(DEFAULT_MEDIA_OUTPUT):
            img_path = os.path.join(DEFAULT_MEDIA_OUTPUT, image)
            _, img_extension = os.path.splitext(img_path)
            img_valid = False
            if(img_extension == ".emf" or img_extension == ".wmf"):
                img_path = convert_anymf_to_png(img_path, img_extension)
                img_extension = ".png"
                img_valid = True
            if(img_extension.lower() in [".jpeg",".jpg",".png"]):
                img_valid = True

            # If image exists, is valid and is not an uuid'd image
            if os.path.isfile(img_path) and img_valid and "image" in image:
                new_img_name = str(uuid.uuid4()) + img_extension
                new_img_path = os.path.join(OBSIDIAN_ATTACHMENTS, new_img_name)
                shutil.move(img_path, new_img_path)
                replace_image_integration_in_file(image, md_file_path, new_img_name)

        shutil.rmtree(DEFAULT_MEDIA_OUTPUT) # Delete the media folder created by pandoc


def convert_anymf_to_png(img_path, img_extension):
    print("Converting " + img_path + " to png")
    print("To : png")
    try:
        cmd_str = "\"" + LIBREOFFICE_EXECUTABLE + "\" --headless --convert-to png \"" + img_path + "\" --outdir \"" + DEFAULT_MEDIA_OUTPUT + "\""
        subprocess.run(cmd_str, shell=True)
    except Exception as e:
        print("Error when converting " + img_path + " to png : " + e)

    os.remove(img_path)
    new_png_image_path = img_path.replace(img_extension, ".png")
    im = Image.open(new_png_image_path)
    im = trim(im)
    im.save(new_png_image_path)
    return new_png_image_path


def convert_doc_to_docx(input_file_path):
    print("Converting " + input_file_path)
    print("To : docx")
    current_dir = os.path.dirname(input_file_path)
    try:
        cmd_str = "\"" + LIBREOFFICE_EXECUTABLE + "\" --headless --convert-to docx \"" + input_file_path + "\"  --outdir \"" + current_dir + "\""
        subprocess.run(cmd_str, shell=True)
    except Exception as e:
        print("Error when converting " + input_file_path + " to doc : " + e)
    
    output_file_path = re.sub("\.doc$", ".docx", input_file_path) # Replace the filename from .docx to .md
    return output_file_path


def convert_pp_to_pdf(input_file_path):
    print("Converting " + input_file_path)
    print("To : pdf")
    current_dir = os.path.dirname(input_file_path)
    try:
        cmd_str = "\"" + LIBREOFFICE_EXECUTABLE + "\" --headless --convert-to pdf \"" + input_file_path + "\"  --outdir \"" + current_dir + "\""
        subprocess.run(cmd_str, shell=True)
    except Exception as e:
        print("Error when converting " + input_file_path + " to doc : " + e)

    return re.sub("\.pptx?$", ".pdf", input_file_path) # Replace the filename from .pptx to .pdf

def trim(im):
    bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return im.crop(bbox)


def replace_image_integration_in_file(image_name, file, new_img_name):
    """ Look for any file that matches an image tag and replace it with the new image name in image obsidian integration ![[123.png]]"""
    with open (file, 'r', encoding='utf-8') as opened_file:
        content = opened_file.read()
        content_new = re.sub('\<img src=\".*' + image_name + '\".*\>', "![[" + new_img_name + "]]", content) # markdown_mmd
        # content_new = re.sub('\!\[\]\(.*' + image_name + '\)({.*})?', "![[" + new_img_name + "]]", content) # markdown
    with open(file, 'wb') as opened_file:
        opened_file.write(content_new.encode('utf8', 'ignore'))
        
def clean_file(file):
    """ Clean file of other doc(x) artefacts"""
    with open (file, 'r', encoding='utf-8') as opened_file:
        content = opened_file.read()
        content_new = re.sub("\[(.*)\]{.underline}", r'<u>\1</u>', content) # Replace all [text]{.underline} with <u>text</u>
        
    with open(file, 'wb') as opened_file:
        opened_file.write(content_new.encode('utf8', 'ignore'))

def create_markdown_link_file(file_path, file_content):
    with open(file_path, 'wb') as new_file:
        new_file.write(file_content.encode('utf8', 'ignore'))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description = 'Convert docx to md, for obsidian import')

    parser.add_argument("input_directory", help="The input directory with docx files")
    parser.add_argument("obsidian_attachment_directory", help="Your obsidian's attachment directory")
    parser.add_argument("obsidian_output_directory", help="Your obsidian's output directory")
    parser.add_argument("--libreoffice", required=False, help="The path to your libreoffice executable", default="C:\Program Files\LibreOffice\program\soffice.exe")
    parser.add_argument("--additional_files", required=False, help="A comma separated list of other files that can be copied to the vault (ex: .pdf,.xlsx)")
    parser.add_argument("-r", "--recursive", help="Runs recursively", action='store_true')
    parser.add_argument("-p", "--powerpoint_to_pdf", help="Converts all powerpoints to a pdf file and copy the pdf file", action='store_true')
    parser.add_argument("--linked_files", required=False, help="A comma separated list of other files that will be linked into a new markdown file (ex: .xlsx,.vba). Files need to be part of the additional_files list.")

    args = parser.parse_args()
    OBSIDIAN_ATTACHMENTS = os.path.abspath(args.obsidian_attachment_directory)
    LIBREOFFICE_EXECUTABLE = args.libreoffice
    POWERPOINT_TO_PDF = args.powerpoint_to_pdf
    
    if(args.linked_files == "all"):
        LINK_ALL_FILES = True
    elif(args.linked_files):
        LINK_TO_UNSUPPORTED_FILES = args.linked_files.split(",")

    if(args.additional_files == "all"):
        COPY_ALL_FILES = True
    elif(args.additional_files):
        FILES_TO_COPY = args.additional_files.split(",")

    obsidian_output = os.path.abspath(args.obsidian_output_directory)
    initial_directory = os.path.abspath(args.input_directory)
    
    try:
        if(initial_directory is None or not os.path.isdir(initial_directory)):
            raise Exception("Directory argument cannot be empty or directory doesn't exist")
    
        if(args.recursive):
            for (root,dirs,files) in os.walk(initial_directory, topdown=True):
                relpath_to_deepdir = root.replace(initial_directory, "") # relative path to add to obsidian
                output_dir_path = os.path.join(obsidian_output, relpath_to_deepdir[1:]) # remove 1st \ to relpath to fix issue path issue
                convert_files(output_dir_path, root, files)
        else:
            files = os.listdir(initial_directory)
            convert_files(obsidian_output, initial_directory, files)

    except KeyboardInterrupt:
        print("Done")
        sys.exit(0)
    except Exception as e:
        print(e)
        sys.exit(1)

    sys.exit(0)
