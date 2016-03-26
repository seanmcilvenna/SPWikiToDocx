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
        describe: 'The file to save the output to',
        default: 'output.docx'
    })
    .option('images', {
        alias: 'i',
        describe: 'The library that all images are stored in for the content in the library'
    })
    .argv;

var client = new SP.RestService(argv.site);
var library = client.list(argv.library);
var images;

if (argv.images) {
    images = client.list(argv.images);
}

function getImage(list, imageUrl) {
    var deferred = Q.defer();

    var options = {
        url: imageUrl,
        headers: {
            'Cookie': 'FedAuth=' + list.service.FedAuth + '; rtFa=' + list.service.rtFa,
        },
        encoding: null
    };

    request(options, function(error, response, body) {
        if (error) {
            deferred.reject(error);
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

function replaceImages(allHtmlContent, allImages) {
    if (!allImages) {
        return Q.resolve(allHtmlContent);
    }

    var deferred = Q.defer();
    var parser = new xml2js.Parser();
    var builder = new xml2js.Builder();

    var fixImageTags = function(current) {
        if (current instanceof Array) {
            for (var i in current) {
                fixImageTags(current[i]);
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

                        for (var x in allImages) {
                            if (allImages[x].path == img['$']['src']) {
                                img['$']['src'] = 'data:' + allImages[x].type + ';base64,' + allImages[x].content;
                            }
                        }
                    }
                } else {
                    fixImageTags(current[i]);
                }
            }
        }
    }

    parser.parseString(allHtmlContent, function(err, result) {
        if (err) {
            deferred.reject(err);
        } else {
            fixImageTags(result);
            var xml = builder.buildObject(result);
            deferred.resolve(xml);
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

        var allImages = {};

        /*
        client.metadata(function(err, data) {
            console.log('test');
        });
        */

        getLibraryItems(images)
            .then(function(data) {
                if (data) {
                    var imagePromises = [];

                    for (var i in data.results) {
                        if (data.results[i].ContentType == 'Image') {
                            imagePromises.push(
                                getImage(images, data.results[i].__metadata.media_src)
                            );
                        }
                    }

                    return Q.all(imagePromises);
                }
            })
            .then(function(imageDatas) {
                allImages = imageDatas;
                return getLibraryItems(library);
            })
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

                return replaceImages(allHtmlContent, allImages);
            })
            .then(function(results) {
                var docx = HtmlDocx.asBlob(results);
                //fs.writeFile('test.html', results);
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