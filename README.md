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
npm install sp-wiki-to-docx
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