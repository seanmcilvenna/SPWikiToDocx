# SPWikiToDocx
Convert a sharepoint wiki library into a DOCX file

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

After installing, because of an issue with the "sharepoint" dependency, you have to CD to the directory that the sharepoint module is installed and forcefully install xml2js@0.1.14

```
cd "\Program Files\nodejs\node_modules\sp-wiki-to-docx\node_modules\sharepoint"
npm install xml2js@0.1.14
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