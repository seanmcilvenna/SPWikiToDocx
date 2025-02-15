# SPWikiToDocx
Convert a sharepoint wiki library into a DOCX file.
The tool reads each Wiki page in a given SP library and combines all of the contents of the Wiki pages into a single HTML chunk. html-docx-js is used to convert that one large HTML chunk into a DOCX file.
The tool attempts to retrieve each of the images referenced by the Wiki pages and embeds the images as base64 data in the HTML files ("data:image/png;base64,XXXXX") so that the images are visible in the output.

# Dependencies
* html-docx-js
* q
* readline
* request
* sharepoint
* xml2js
* yargs

# Installation
```
npm install sp-wiki-to-docx -g
```

# Usage
```
sp-wiki-to-docx -s SITE_URL -l LIBRARY_NAME -u USERNAME -o output.docx
```

**Options:**

| Long | Short | Description | Required? |
| ---- | ----- | ----------- | --------- |
| --site | -s | The site URL | Yes |
| --library | -l | The name of the library (no spaces)\nex: "GeneralGuides" | Yes | 
| --username | -u | Sharepoint online username\nEx: my@email.com | Yes |
| --password | -p | Sharepoint online password | No, will prompt if not specified on command line |
| --output | -o | The file to save the output to (default: "output.docx") | No |
| --combinedHtml | -c | The file to save the combined HTML output to\nex: output.html | No |

## ☕ Support My Work

Find this useful and want to show appreciation??  

[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-Support%20My%20Work-orange?style=flat&logo=buy-me-a-coffee)](https://buymeacoffee.com/seanmcilvenna)

Your support helps keep this project alive and motivates me to continue improving it! 🚀
