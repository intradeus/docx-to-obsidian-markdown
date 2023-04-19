# Docx to obsidian

This repo contains a script that can batch import docx files into your obsidian vault.

## How it works :
1) All files in the input folder will be converted to an obsidian-flavoured markdown, thanks to pandoc and some custom code.
2) Images extracted by pandoc will be renamed to a random name (2acfacdf-008d-4f51-9e4e-041720526a04.png) to avoid conflicts
   1) .emf, .wmf (word images format) will be converted to PNG, thanks to libreoffice
3) All image backlinks in your documents will be updated to the new UUID provided and transformed to an obsidian image backlink (![[image.png]])
   1) All existing relative links converted by pandoc (![](C:\path\to\image.png)) will therefore be integrated according to the obsidian standard 

Note: The script is not recursive, it will only import the files from the input directory to the output directory.

## Speeds
Speeds depends on a lot of factor, including if your images are exported as .emf a lot. What takes the most time is the conversion from .emf to .png.
I've had runs where 10 docx files would take about 15 seconds, but some others would take 45 seconds.

## Requirements

1) You must have [Pandoc](https://pandoc.org/installing.html) installed
2) You must have [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/) installed. This is because we need to use LibreOffice to convert .emf and .wmf image files to png.
3) You must install python requirements : 
```
python -m pip install -r requirements.txt 
```


## Test

Once requirements are installed, you can test the script yourself with the test folder provided by running the following command: 
```
python converter.py ./input_test ./obsidian_vault/attachments ./obsidian_vault 
```

## Run 
You can import in your vault using:
```
python converter.py <path-to-input> <path-to-obsidian-attachment> <path-to-obsidian-output>
```

If you don't have LibreOffice installed in the default path (C:\Program Files\LibreOffice\program\soffice.exe) you will need to add the following argument : 
```
--libreoffice <path-to-soffice.exe>
```

Exemple : 
```
python converter.py "C:\Documents\Old Word Docs\Biologie" "C:\Documents\Obsidian Vault\Attachments" "C:\Documents\Obsidian Vault\Biologie"
```

## Helping 
I used this script to import many docx into my obsidian vault and was able to fix a few errors i've encountered, but if you notice new errors OR would like to add more modifications after the conversion, open an issue or a PR :)

## Roadmap
The next implementation I wanna add are :
- a recursive argument, to recursively import files.


