## Batch converter for docx to obsidian flavoured markdown

## Requirements

1) You must have pandoc installed
2) You must have LibreOffice installed. This is because we need to use LibreOffice to convert .emf and .wmf image files to png.
3) You must install python requirements : 
```
python -m pip install -r requirements.txt 
```

## Test

You can test with the test folder provided by running the following command: 
```
python converter.py ./input_test ./obsidian_vault/attachments ./obsidian_vault 
```

## Run 
You can run to add in your vault with : 
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

## How it works :
1) All files in `C:\Documents\Old Word Docs\Biologie` will be converted to an obsidian-flavoured markdown, thanks to pandoc
2) Images will be renamed to a UUID
   1) png, jpeg etc. are not touched
   2) .emf, .wmf (word images format) will be converted to PNG, thanks to libreoffice
3) All image backlinks in your documents will be updated to the new UUID provided, and transformed to an obsidian image backlink (![[image.png]])
4) 