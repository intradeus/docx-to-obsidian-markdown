import os, sys, argparse
import subprocess
import uuid
import re
from PIL import Image, ImageChops
import shutil
from pathlib import Path

OBSIDIAN_ATTACHMENTS = ""
OBSIDIAN_OUTPUT = ""
LIBREOFFICE_EXECUTABLE = ""
DEFAULT_MEDIA_OUTPUT = ""
FILES_TO_COPY = []

def convert_directories(dir):
    """ Recursive directory batch conversion/importation"""
    if(dir is None):
        raise Exception("Directory argument cannot be empty")
    
    
    for (root,_,files) in os.walk(dir, topdown=True):
        for file in files:
            global OBSIDIAN_OUTPUT
            _, file_extension = os.path.splitext(file)
            input_file_path = os.path.join(root, file)
            relpath_to_deepdir = root.replace(dir, "") # relative path to add to obsidian
            output_dir_path = os.path.join(OBSIDIAN_OUTPUT, relpath_to_deepdir[1:]) # remove 1st \ to relpath to fix issue path issue
            output_file_path = os.path.join(output_dir_path, file)
            # If file is docx, convert it
            if os.path.isfile(input_file_path) and file_extension in [".docx", ".doc"]:
                output_file_path = re.sub("\.docx?", ".md", output_file_path) # Replace the filepath from .doc(x) to .md
                current_dir = os.path.dirname(input_file_path)
                Path(output_dir_path).mkdir(parents=True, exist_ok=True) # Create necessary directories
                print(input_file_path)
                print(output_file_path)
                print(current_dir)
                print("---")
                run_conversion(current_dir, input_file_path, output_file_path)

            # If file is part of the files-to-copy arguments, just copy it to the obsidian vault
            if os.path.isfile(input_file_path) and file_extension in FILES_TO_COPY:
                shutil.copyfile(input_file_path, output_file_path)
    
def convert_directory(dir):
    """ Single directory batch conversion/importation"""
    if(dir is None):
        raise Exception("Directory argument cannot be empty")
    
    for filename in os.listdir(dir):
        # checking if it is a file
        _, file_extension = os.path.splitext(filename)
        input_file_path = os.path.join(dir, filename)
        output_file_path = os.path.join(OBSIDIAN_OUTPUT, filename) # Set the output folder as the obsidian output folder

        if os.path.isfile(input_file_path) and file_extension in [".docx", ".doc"]:
            print("Working on " + input_file_path)
            output_file_path = re.sub("\.docx?", ".md", output_file_path) # Replace the filename from .doc(x) to .md
            run_conversion(dir, input_file_path, output_file_path)
        # If file is part of the files-to-copy arguments, just copy it to the obsidian vault
        if os.path.isfile(input_file_path) and file_extension in FILES_TO_COPY:
            shutil.copyfile(input_file_path, output_file_path)


def run_conversion(dir, input_file_path, output_file_path):
    global DEFAULT_MEDIA_OUTPUT
    DEFAULT_MEDIA_OUTPUT = os.path.join(dir, 'media') # Default pandoc output for images is a folder called /media

    # Use pandoc to convert the doc(x) to markdown, using standalone tags and extracting media to our dir folder (it will create a media subfolder)
    try:
        cmd_str = "pandoc -f docx -t markdown -s --wrap=none --extract-media=\""+ dir + "\" \"" + input_file_path + "\" -o \"" + output_file_path + "\""
        subprocess.run(cmd_str, shell=True)
    except Exception as e:
        print("Error while converting from doc(x) to markdown : " + e)
    # Remove any known artefacts from the file
    clean_file(output_file_path)

    clean_images(output_file_path)

    
def clean_file(file):
    """ Clean file of other docx artefacts"""
    with open (file, 'r') as opened_file:
        content = opened_file.read()
        content_new = re.sub("\[(.*)\]{.underline}", r'<u>\1</u>', content) # Replace all [text]{.underline} with just "text"
        
    with open(file, 'w') as opened_file:
        opened_file.write(content_new)

def clean_images(md_file_path):
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

def convert_anymf_to_png(img_path, img_extension):
    print("converting file " + img_path + " to png")
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

def trim(im):
    bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
    diff = ImageChops.difference(im, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        return im.crop(bbox)


def replace_image_integration_in_file(image_name, file, new_img_name):
    """ Look for any file that matches a ![](*image*.*){*} regex and replace it with the new image name in image obsidian integration ![[123.png]]"""
    with open (file, 'r') as opened_file:
        content = opened_file.read()
        content_new = re.sub('\!\[\]\(.*' + image_name + '\)({.*})?', "![[" + new_img_name + "]]", content)
    with open(file, 'w') as opened_file:
        opened_file.write(content_new)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description = 'Convert docx to md, for obsidian import')

    parser.add_argument("input_directory", help="The input directory with docx files")
    parser.add_argument("obsidian_attachment_directory", help="Your obsidian's attachment directory")
    parser.add_argument("obsidian_output_directory", help="Your obsidian's output directory")
    parser.add_argument("--libreoffice", required=False, help="The path to your libreoffice executable", default="C:\Program Files\LibreOffice\program\soffice.exe")
    parser.add_argument("--files_to_copy",required=False, help="A comma separated list of other files that can be copied to the vault", default=".pdf")
    parser.add_argument("-r", "--recursive", help="Runs recursively", action='store_true')

    args = parser.parse_args()
    OBSIDIAN_ATTACHMENTS = os.path.abspath(args.obsidian_attachment_directory)
    OBSIDIAN_OUTPUT = os.path.abspath(args.obsidian_output_directory)
    LIBREOFFICE_EXECUTABLE = args.libreoffice
    FILES_TO_COPY = args.files_to_copy.split(",")


    try:
        if(args.recursive):
            convert_directories(os.path.abspath(args.input_directory))
        else:
            convert_directory(args.input_directory)
    except KeyboardInterrupt:
        print("Done")
        sys.exit(0)
    except Exception as e:
        print(e)
        sys.exit(1)

    sys.exit(0)