import os, sys, argparse
import subprocess
import uuid
import re
from PIL import Image, ImageChops
import shutil

OBSIDIAN_ATTACHMENTS = ""
OBSIDIAN_OUTPUT = ""
LIBREOFFICE_EXECUTABLE = ""
DEFAULT_MEDIA_OUTPUT = ""


def convert_directory(dir):
    if(dir is None):
        raise Exception("Directory argument cannot be empty")
    
    for filename in os.listdir(dir):
        # checking if it is a file
        _, file_extension = os.path.splitext(filename)
        if os.path.isfile(os.path.join(dir, filename)) and file_extension in [".docx", ".doc"]:
            print("Working on " + filename)
            run_conversion(dir, filename)


def run_conversion(dir, filename):
    input_file_path = os.path.join(dir, filename)
    output_filename = re.sub(".docx?", ".md", filename) # Replace the filename from .doc(x) to .md
    output_file_path = os.path.join(OBSIDIAN_OUTPUT, output_filename) # Set the output folder as the obsidian output folder

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
    args = parser.parse_args()
    OBSIDIAN_ATTACHMENTS = args.obsidian_attachment_directory
    OBSIDIAN_OUTPUT = args.obsidian_output_directory
    LIBREOFFICE_EXECUTABLE = args.libreoffice

    try:
        convert_directory(args.input_directory)
    except KeyboardInterrupt:
        print("Done")
        sys.exit(0)
    except Exception as e:
        print(e)
        sys.exit(1)

    sys.exit(0)