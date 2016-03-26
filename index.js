#! /usr/bin/env node

var HtmlDocx = require('html-docx-js');
var SP = require('sharepoint');
var Q = require('q');
var fs = require('fs');
var readline = require('readline');
var request = require('request');
var xml2js = require('xml2js');
var url = require('url');
var argv = require('yargs')
    .option('site', {
        alias: 's',
        describe: 'The site URL',
        required: true
    })
    .option('library', {
        alias: 'l',
        describe: 'The name of the library (no spaces)\nex: "GeneralGuides"',
        required: true
    })
    .option('username', {
        alias: 'u',
        describe: 'Sharepoint online username\nex: sean.mcilvenna@lantanagroup.com',
        required: true
    })
    .option('password', {
        alias: 'p',
        describe: 'Sharepoint online password'
    })
    .option('output', {
        alias: 'o',
        describe: 'The file to save the DOCX output to',
        default: 'output.docx'
    })
    .option('combinedHtml', {
        alias: 'c',
        describe: 'The file to save the combined HTML output to'
    })
    .argv;

var client = new SP.RestService(argv.site);
var library = client.list(argv.library);

function getImage(imageUrl, fromSharepoint) {
    var deferred = Q.defer();

    var requestUrl = imageUrl;

    if (requestUrl.indexOf('http://') < 0 && requestUrl.indexOf('https://') < 0) {
        requestUrl = library.service.protocol + '//' + library.service.host + requestUrl;
    }

    var options = {
        url: requestUrl,
        headers: {},
        encoding: null
    };

    if (fromSharepoint) {
        options.headers['Cookie'] = 'FedAuth=' + library.service.FedAuth + '; rtFa=' + library.service.rtFa;
    }

    request(options, function(error, response, body) {
        if (error) {
            deferred.reject(error);
        } else if (response.statusCode != 200) {
            deferred.reject(body.toString('utf8'));
        } else {
            var imageUrlInfo = url.parse(imageUrl);
            deferred.resolve({
                path: imageUrlInfo.path,
                type: response.headers['content-type'],
                content: body.toString('base64')
            });
        }
    });

    return deferred.promise;
};

function getLibraryItem(list, id) {
    var deferred = Q.defer();

    list.get(id, function(err, data) {
        if (err) {
            deferred.reject(err);
        } else {
            deferred.resolve(data);
        }
    });

    return deferred.promise;
};

function getLibraryItems(list) {
    if (!list) {
        return Q.resolve();
    }

    var deferred = Q.defer();

    list.get(function(err, data) {
        if (err) {
            deferred.reject(err);
        } else {
            deferred.resolve(data);
        }
    });

    return deferred.promise;
};

function replaceImage(imgTag) {
    var deferred = Q.defer();
    var imgSrc = imgTag['$']['src'];
    var getImgPromise;

    if (imgSrc.indexOf(library.service.path) == 0) {
        getImgPromise = getImage(imgSrc, true);
    } else {
        getImgPromise = getImage(imgSrc, false);
    }

    getImgPromise
        .then(function(imgData) {
            imgTag['$']['src'] = 'data:' + imgData.type + ';base64,' + imgData.content;
        })
        .catch(function() {
            console.log('Could not retrieve image: ' + imgSrc);
        })
        .done(function() {
            deferred.resolve();
        });

    return deferred.promise;
};

function replaceImages(allHtmlContent) {
    var deferred = Q.defer();
    var parser = new xml2js.Parser();
    var builder = new xml2js.Builder();

    var findImageTags = function(current) {
        var imgTags = [];

        if (current instanceof Array) {
            for (var i in current) {
                imgTags = imgTags.concat(findImageTags(current[i]));
            }
        } else if (typeof current == 'object') {
            for (var i in current) {
                if (i == 'img') {
                    if (current[i].length == 0) {
                        continue;
                    }

                    for (var y in current[i]) {
                        var img = current[i][y];

                        if (!img['$'] || !img['$']['src']) {
                            continue;
                        }

                        imgTags.push(img);
                    }
                } else {
                    imgTags = imgTags.concat(findImageTags(current[i]));
                }
            }
        }

        return imgTags;
    };

    parser.parseString(allHtmlContent, function(err, result) {
        if (err) {
            deferred.reject(err);
        } else {
            var imgTags = findImageTags(result);
            var replaceImagePromises = [];

            for (var i in imgTags) {
                var imgTag = imgTags[i];
                var replaceImagePromise = replaceImage(imgTag);
                replaceImagePromises.push(replaceImagePromise);
            }

            Q.all(replaceImagePromises)
                .then(function() {
                    var xml = builder.buildObject(result);
                    deferred.resolve(xml);
                })
                .catch(function(err) {
                    deferred.reject(err);
                });
        }
    });

    return deferred.promise;
};

function doWork() {
    client.signin(argv.username, argv.password, function(err, data) {
        // check for errors during login, e.g. invalid credentials
        if (err) {
            console.log("Error: ", err);
            process.exit(1);
            return;
        }

        /*
        client.metadata(function(err, data) {
            console.log('test');
        });
        */

        getLibraryItems(library)
            .then(function(libraryData) {
                var itemPromises = [];

                for (var i in libraryData.results) {
                    var itemPromise = getLibraryItem(library, libraryData.results[i].Id);
                    itemPromises.push(itemPromise);
                }

                return Q.all(itemPromises);
            })
            .then(function(results) {
                var allHtmlContent = '<div>';

                for (var i in results) {
                    var result = results[i];

                    if (result.ContentType != 'Wiki Page') {
                        continue;
                    }

                    var pageName = result.Name.substring(0, result.Name.lastIndexOf('.aspx'));
                    allHtmlContent += '<div><h1 style="text-decoration: underline">' + pageName + '</h1>' + result.WikiContent + '</div>';
                }

                allHtmlContent += '</div>';

                allHtmlContent = allHtmlContent.replace(/<br>/g, '<br/>');

                return replaceImages(allHtmlContent);
            })
            .then(function(results) {
                var docx = HtmlDocx.asBlob(results);

                if (argv.combinedHtml) {
                    fs.writeFile(argv.combinedHtml, results);
                }

                fs.writeFile(argv.output, docx, function(err) {
                    if (err) throw err;
                    console.log('Done');
                    process.exit(0);
                });
            })
            .catch(function(err) {
                console.log('Error: ' + err);
                process.exit(1);
            });
    });
};

if (!argv.password) {
    var i = readline.createInterface(process.stdin, process.stdout, null);
    i.question('Please enter your password\n', function(password) {
        argv.password = password;
        i.close();
        doWork();
    });
} else {
    doWork();
}