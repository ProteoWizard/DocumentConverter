var mammoth = require("./index");
var fs = require('fs');
var path = require('path');
var beautify = require("js-beautify")
var _ = require('underscore');
var exec = require('child_process');
var crypto = require('crypto')

//currently lives in skyline-resources
var imageConverterPath = `${__dirname}/skyline-resources/ImageConverter-net8.0-windows/ImageConverter.exe`;

var htmlPrettyPrintOptions = options = { wrap_line_length: 150};
var defaultSharedImageThreshold = 3;

var imageContentTypeExtensionMap = {
    'image/png': '.png',
    'image/x-emf': '.emf',
    'image/x-wmf': '.wmf',
    'image/jpeg': '.jpeg'
}

var locales = ["en", "ja", "zh-CHS"]

var ignoredWarnings = [
    "Unrecognised paragraph style: 'List Paragraph' (Style ID: ListParagraph)", //verified no additional formatting needed
    "Image of type image/x-emf is unlikely to display in web browsers", //post processor conversion
    "Image of type image/x-wmf is unlikely to display in web browsers", //post processor conversion
    "Unrecognised run style: 'Internet Link' (Style ID: InternetLink)", //verified no additional formatting needed
    "Unrecognised paragraph style: 'No Spacing' (Style ID: NoSpacing)", // ignoring style
    "Unrecognised paragraph style: 'Body Text' (Style ID: BodyText)", // ignoring style
    "Unrecognised paragraph style: 'Normal (Web)' (Style ID: NormalWeb)", // ignoring style
    "Unrecognised paragraph style: 'Default' (Style ID: Default)", // ignoring style

    // "An unrecognised element was ignored: {urn:schemas-microsoft-com:office:office}lock", //verified no additional formatting needed
    // "A v:imagedata element without a relationship ID was ignored", //verified no additional formatting needed
    // "An unrecognised element was ignored: v:fill", //documented but no solution possible
    // "An unrecognised element was ignored: v:stroke", //documented but no solution possible
    // "An unrecognised element was ignored: v:path", //documented but no solution possible
    // "An unrecognised element was ignored: v:oval", //documented but no solution possible
    // "An unrecognised element was ignored: office-word:anchorlock", //documented but no solution possible
    // "An unrecognised element was ignored: {urn:schemas-microsoft-com:office:office}OLEObject", //documented but no solution possible
    // "An unrecognised element was ignored: {http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara" //documented but no solution possible
]

var coverShotNameMap = {
    'Skyline Absolute Quantification': 'AbsoluteQuant',
    'Skyline Advanced Peak Picking': 'PeakPicking',
    'Skyline Audit Logging': 'AuditLog',
    'Skyline Collision Energy Optimization': 'OptimizeCE',
    'Skyline Custom Reports': 'CustomReports',
    'Skyline Data Independent Acquisition': 'DIA',
    'Skyline DIA PASEF': 'DIA-PASEF',
    'Skyline DIA QE': 'DIA-QE',
    'Skyline DIA TTOF': 'DIA-TTOF',
    'Skyline DIA Umpire TTOF': 'DIA-Umpire-TTOF',
    'Skyline Existing and Quantitative Experiments': 'ExistingQuant',
    'Skyline Hi-Res Metabolomics': 'HiResMetabolomics',
    'Skyline Importing Assay Libraries': 'ImportingAssayLibraries',
    'Skyline Importing Integration Boundaries': 'ImportingIntegrationBoundaries',
    'Skyline Ion Mobility Spectrum Filtering': 'IMSFiltering',
    'Skyline iRT Retention Time Prediction': 'iRT',
    'Skyline MS1 DDA Search': 'DDASearch',
    'Skyline MS1 Filtering': 'MS1Filtering',
    'Skyline PRM': 'PRM',
    'Skyline PRM Orbitrap': 'PRMOrbitrap',
    'Skyline PRM Orbitrap-PRBB-format': 'PRMOrbitrap-PRBB',
    'Skyline Processing Grouped Study Data': 'GroupedStudies',
    'Skyline Small Molecule Method Dev and CE Opt': 'SmallMoleculeMethodDevCEOpt',
    'Skyline Small Molecule Multidimensional Spectral Libraries': 'SmallMoleculeIMSLibraries',
    'Skyline Small Molecule Quantification': 'SmallMoleculeQuantification',
    'Skyline Small Molecule Targets': 'SmallMolecule',
    'Skyline Spectral Library Explorer': 'LibraryExplorer',
    'Skyline Targeted Method Editing': 'MethodEdit',
    'Skyline Targeted Method Refinement': 'MethodRefine'
}

//stateful
var imagesToConvert = [];
var imageHashCount = {}
var sharedImageNames = {}
var sharedImageCount = 0;

var debugOptions = {
    sourcePath: "C:\\Users\\Eduardo\\Documents\\TutorialSandbox",
    existingSharedPath: "C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials/shared",
}
//main(debugOptions);

function main(options){
    options = setDefaultOptions(options)

    createSharedFolder(options).then(function() {
        convertAllDocx(options, storeSharedImages).then(function() {
            convertAllDocx(options, convertDocument).then(function () {
                rearrangeDirectoryStructure(options);
                postProcess(options);
            })
        })
    })
}

function postProcess(options){
    options = setDefaultOptions(options)
    var files = fs.readdirSync(options.sourcePath, { recursive: true });
    files.filter(
        function(file) {
            return file.endsWith(".html");
        }).forEach(function(sourceDocument) {
            var filePath = `${options.sourcePath}/${sourceDocument}`;
            fs.readFile(filePath, "utf8", (err, data) => {
                if (err) throw err;
                var html = data;
                var processCounter = 0

                while(true){
                    processCounter++;
                    var newHtml = postProcessHtml(html, filePath)
                    if(newHtml === html){
                        break;
                    }
                    if(processCounter >= 30){
                        console.log(`Html processing reached limit of ${processCounter} iterations for document ${sourceDocument}`)
                        throw `Html processing reached limit of ${processCounter} iterations for document ${sourceDocument}`;
                    }
                    html = newHtml;
                }

                fs.writeFile(filePath, html, "utf8", (err) => {
                    if (err) throw err;

                    console.log(`Successfully processed document ${sourceDocument} in ${processCounter} iterations`);
                });
            });
        });
}

function postProcessHtml(html, htmlPath){
    //replace strong tag with b, em tag with i
    html = html.replace(/<strong>(.*?)<\/strong>/gs, '<b>$1</b>');
    html = html.replace(/<em>(.*?)<\/em>/gs, '<i>$1</i>');

    //Remove b/i tags that only contain a colon or period or comma
    html = html.replace(/<b>:<\/b>/gs, ':');
    html = html.replace(/<b>\.<\/b>/gs, '.');
    html = html.replace(/<b>,<\/b>/gs, ',');
    html = html.replace(/<i>:<\/i>/gs, ':');
    html = html.replace(/<i>\.<\/i>/gs, '.');
    html = html.replace(/<i>,<\/i>/gs, ',');

    //remove b/i tags with whitespace only but preserve whitespace
    html = html.replace(/<b>(\s*)<\/b>/gs, '$1');
    html = html.replace(/<i>(\s*)<\/i>/gs, '$1');

    //move periods or commas or colons outside of b or i tags
    html = html.replace(/(<b>.*?)(\.)<\/b>/gs, '$1</b>$2');
    html = html.replace(/(<b>.*?)(,)<\/b>/gs, '$1</b>$2');
    html = html.replace(/(<b>.*?)(:)<\/b>/gs, '$1</b>$2');
    html = html.replace(/(<i>.*?)(\.)<\/i>/gs, '$1</i>$2');
    html = html.replace(/(<i>.*?)(,)<\/i>/gs, '$1</i>$2');
    html = html.replace(/(<i>.*?)(:)<\/i>/gs, '$1</i>$2');

    //move leading whitespace outside of b or i tags
    html = html.replace(/<b>\s+(.*)<\/b>/gs, ' <b>$1</b>');
    html = html.replace(/<i>\s+(.*)<\/i>/gs, ' <i>$1</i>');

    //move trailing whitespace outside of b or i tags
    html = html.replace(/<b>(.*?)\s+<\/b>/gs, '<b>$1</b> ');
    html = html.replace(/<i>(.*?)\s+<\/i>/gs, '<i>$1</i> ');

    //remove b/i tags which are empty
    html = html.replace(/<b><\/b>/gs, '');
    html = html.replace(/<i><\/i>/gs, '');

    //remove empty lines before p closing tags
    html = html.replace(/^\s*$(?=\n\s*<\/p>\s*)/gm, '');

    //remove empty anchor tags
    html = html.replace(/<a\s+[^>]*><\/a>/gs, '');

    //add 0 wrapper to single integers
    html = html.replace(/(s-)(\d)(\.png)/gs, (match, prefix, number, suffix) => {
        const paddedNumber = number.padStart(2, '0');
        const oldImageName = `${prefix}${number}${suffix}`
        const newImageName = `${prefix}${paddedNumber}${suffix}`
        const imageDir = path.dirname(htmlPath)
        if(fs.existsSync(`${imageDir}/${oldImageName}`)){
            fs.renameSync(`${imageDir}/${oldImageName}`, `${imageDir}/${newImageName}`);
        }
        return newImageName;
    });

    return html;
}

function setDefaultOptions(options){
    if(!options.subPaths) options.subPaths = [];
    if(!options.htmlName) options.htmlName = "index.html";
    if(!options.sharedImageThreshold) options.sharedImageThreshold = defaultSharedImageThreshold;
    if(!options.outputPath) options.outputPath = options.sourcePath + "/output";
    if(!options.invariantFileNames) options.invariantFileNames = [ "invariant", "invariantDraft" ]
    if(options.regenerateImages === undefined) options.regenerateImages = true;
    if(options.postProcessOnly === undefined) options.postProcessOnly = false;
    return options;
}

function convertAllDocx(options, convertFunction){
    var documentConversionPromises = [];
    documentConversionPromises.push(...convertDirectory(options, options.sourcePath, options.outputPath, convertFunction))
    options.subPaths.forEach(function(subPath){
        documentConversionPromises.push(...convertDirectory(options,`${options.sourcePath}/${subPath}`, `${options.outputPath}/${subPath}`, convertFunction));
    })

    //converting images at the time of document generation causes file conflict issues
    //to avoid this we convert all the images once the document generation is complete
    return Promise.all(documentConversionPromises)
        .then(function(results){
            var imageConversionPromises = [];
            imagesToConvert.forEach(function(image){
                imageConversionPromises.push(convertToPng(image));
            })
            return Promise.all(imageConversionPromises);
        })
}


function convertDirectory(options, currentSourcePath, currentOutputPath, convertFunction){
    var documentConversionPromises = [];
    if (!fs.existsSync(currentSourcePath)) return [];
    if (!fs.existsSync(currentOutputPath)) fs.mkdirSync(currentOutputPath);
    var files = fs.readdirSync(currentSourcePath, { recursive: false });
    files.filter(
        function(file) {
            return file.endsWith(".docx") && !file.includes("~$");
        }).forEach(function(sourceDocument) {
            documentConversionPromises.push(convertFunction(options, currentSourcePath, currentOutputPath, sourceDocument));
        });
    return documentConversionPromises;
}

function storeSharedImages(options, currentSourcePath, currentOutputPath, sourceDocument) {
    if(options.existingSharedPath || options.sharedImageNameOverride){
        return Promise.resolve();
    }
    var mammothOptions = {
        convertImage: mammoth.images.imgElement(function(image) {
            var imageCount;
            var hash;
            return image.readAsBase64String().then(function(base64Image) {
                hash = crypto.createHash('md5').update(base64Image).digest('hex');
                imageCount = imageHashCount[hash] ? imageHashCount[hash] + 1 : 1;
                imageHashCount[hash] = imageCount;
                return image;
            }).then(function(image) {
                image.readAsBuffer().then(function(imageBuffer) {
                    if((image.contentType === "image/png" ||  image.contentType === "image/jpeg") && imageCount === options.sharedImageThreshold){
                        var imageFileName = `shared-image-${sharedImageCount}${imageContentTypeExtensionMap[image.contentType]}`
                        sharedImageCount = sharedImageCount + 1;
                        sharedImageNames[hash] = imageFileName;
                        fs.writeFileSync(`${options.outputPath}/shared/${imageFileName}`, imageBuffer);
                    }
                    return {};
                })
            })
        })
    };
    return mammoth.convertToHtml({path: currentSourcePath + "/" + sourceDocument}, mammothOptions);
}

function convertDocument(options, currentSourcePath, currentOutputPath, sourceDocument) {
    var documentName = path.parse(sourceDocument).name;
    var outputFileName = options.invariantFileNames.includes(documentName) ? `${documentName}.html` : options.htmlName;
    var outputFolder = currentOutputPath + "/" + documentName;
    var outputFile = outputFolder + "/" + outputFileName;
    var imageCounter = 1;
    if (!fs.existsSync(outputFolder)) fs.mkdirSync(outputFolder);

    var mammothOptions = {
        styleMap: [
            "p[style-name='Title'] => h1.document-title",
            "p[style-name='Bibliography'] => p.bibliography:fresh",
            "p[style-name='Bibliography1'] => p.bibliography:fresh",
            "p[style-name='Subtitle'] => p.subtitle",
            "r[style-name='Subtle Emphasis'] => b",
        ],
        convertImage: mammoth.images.imgElement(function(image) {
            var hash;
            return image.readAsBase64String().then(function(base64Image) {
                hash = crypto.createHash('md5').update(base64Image).digest('hex');
                return image;
            }).then(function(image) {
                return image.readAsBuffer().then(function(imageBuffer) {
                    var imageLinkName;
                    if(options.sharedImageNameOverride && options.sharedImageNameOverride[hash]){
                        imageLinkName = "../../shared/" + options.sharedImageNameOverride[hash];

                        //write image if it has not been written
                        var sharedImagePath = `${options.outputPath}/shared/${options.sharedImageNameOverride[hash]}`;
                        if(!fs.existsSync(sharedImagePath)){
                            fs.writeFileSync(sharedImagePath, imageBuffer);
                        }
                    }
                    else if(!options.sharedImageNameOverride && sharedImageNames[hash]){
                        imageLinkName = "../../shared/" + sharedImageNames[hash];
                    } else{
                        if (!(image.contentType in imageContentTypeExtensionMap)) {
                            throw "Unknown image type";
                        }
                        var extension = imageContentTypeExtensionMap[image.contentType];
                        var imageName = "s-" +imageCounter++;
                        var imageFileName = imageName +extension;
                        imageLinkName = imageFileName;

                        //rename link to png as we are converting these files
                        if(extension === ".emf" || extension === ".wmf"){
                            imageLinkName = imageName + ".png"
                            if(!options.invariantFileNames.includes(documentName) && options.regenerateImages) {
                                imagesToConvert.push(outputFolder + "/" +imageFileName)
                            }
                        }
                        if(!options.invariantFileNames.includes(documentName) && options.regenerateImages) {
                            fs.writeFileSync(outputFolder + "/" + imageFileName, imageBuffer);
                        }
                    }
                    return {
                        attributes : {
                            src: imageLinkName
                        },
                        anchorAttributes : {
                            name: path.parse(imageLinkName).name
                        }
                    };
                })
            })
        })
    };

    return mammoth.convertToHtml({path: currentSourcePath + "/" + sourceDocument}, mammothOptions)
        .then(function(result) {
            var messages = result.messages.filter(function(message) {
                return !ignoredWarnings.includes(message.message)
            })
            if(messages !== undefined && messages.length > 0){
                console.log("Found warnings/errors converting document " +sourceDocument)
                console.log(messages);
            }
            var skylineJsHtml = `<script src="../../shared/skyline.js" type="text/javascript"></script>`
            var skylineCssHtml = `<link rel="stylesheet" type="text/css" href="../../shared/SkylineStyles.css">`
            var wrappedHtml = `<html><head><meta charset="utf-8">${skylineCssHtml}${skylineJsHtml}</head><body onload="skylineOnload();">${result.value}</body></html>`
            var formattedHtml = formatHtml(wrappedHtml);

            fs.writeFileSync(outputFile, formattedHtml)
        })
        .then(function() {
            //Move docx to output directory as well
            if(options.outputPath === options.sourcePath){
                fs.renameSync(currentSourcePath + "/" + sourceDocument, outputFolder +"/" + sourceDocument)
            } else {
                console.log("Source Path " +currentSourcePath + "/" + sourceDocument)
                fs.copyFileSync(currentSourcePath + "/" + sourceDocument, outputFolder +"/" + sourceDocument)
            }
        });
}

function formatHtml(html){
    var prettyHtml = beautify.html_beautify(html, htmlPrettyPrintOptions);
    prettyHtml = prettyHtml.replace(/(\s*)(<p[^>]*>)(.*?)(<\/p>)/gs, '$1$2$1    $3$1$4');
    return prettyHtml;
}

function convertToPng(sourceFile){
    var outputDir = path.parse(sourceFile).dir
    var outputFile = outputDir + "/" +path.parse(sourceFile).name + ".png";

    var commandArgs = [
        'convertEmfToPng',
        '--input-file', sourceFile,
        '--output-file', outputFile
    ]
    return new Promise((resolve, reject) => {
        exec.execFile(imageConverterPath, commandArgs, function (err, stdout, stderr) {
            if (err) {
                console.log("Could not convert image " +sourceFile)
                if (stderr) console.log(`stderr: ${stderr}`);
                if (stdout) console.log(`stderr: ${stdout}`);
                console.log(`Run command manually: ${imageConverterPath} convertEmfToPng --input-file '${sourceFile}' --output-file '${outputFile}'`);
                reject(new Error("Could not convert image " +sourceFile))
            } else {
                fs.rmSync(sourceFile)
                resolve();
            }
        });
    })
}

function rearrangeDirectoryStructure(options){
    var groupedPaths = createGroupedPaths(options.outputPath, options.subPaths);
    for (var groupedPathKey in groupedPaths) {
        var topLevelPath = `${options.outputPath}/${groupedPathKey}`;
        if (!fs.existsSync(topLevelPath)) fs.mkdirSync(topLevelPath);
        var foundPaths = groupedPaths[groupedPathKey];
        foundPaths.forEach(function (foundPath) {
            if(foundPath.includes("ja")){
                var japanesePath = `${topLevelPath}/ja`;
                moveFilesToDirectory(foundPath, japanesePath);
            } else if(foundPath.includes("zh")){
                var chinesePath = `${topLevelPath}/zh-CHS`;
                moveFilesToDirectory(foundPath, chinesePath);
            } else {
                var englishPath = `${topLevelPath}/en`;
                moveFilesToDirectory(foundPath, englishPath);
            }
        })
    }

    for (var folderName in coverShotNameMap){
        var oldPath = `${options.outputPath}/${folderName}`;
        var newPath = `${options.outputPath}/${coverShotNameMap[folderName]}`
        if(fs.existsSync(oldPath)){
            fs.renameSync(oldPath, newPath)
        }
    }

    options.subPaths.forEach(function(subPath){
        fs.rmSync(`${options.outputPath}/${subPath}`, {recursive: true, force: true,});
    })

}

function moveFilesToDirectory(sourceDirectory, outputDirectory){
    if (!fs.existsSync(outputDirectory)) fs.mkdirSync(outputDirectory);
    fs.readdirSync(sourceDirectory)
        .filter(
            function(file) {
                return !file.includes("~$") && !fs.lstatSync(sourceDirectory + "/" + file).isDirectory();
            }).forEach(function(file) {
            fs.renameSync(sourceDirectory + "/" + file, outputDirectory +"/" + file)
        });
}

function createGroupedPaths(outputPath, subPaths){
    var groupedPaths = {};
    fs.readdirSync(outputPath)
        .forEach(function(file){
            if (fs.lstatSync(outputPath + "/" + file).isDirectory() && file !== "shared" && !subPaths.includes(file)){
                groupedPaths[file] = [outputPath + "/" + file];
            }
        });

    var otherOutputPaths = [];
    subPaths.forEach(function(subPath){
        otherOutputPaths.push(`${outputPath}/${subPath}`);
    })

    otherOutputPaths.forEach(function(otherPath){
        addOtherPath(otherPath);
    });

    return groupedPaths;

    function addOtherPath(path){
        // populate base map
        fs.readdirSync(path)
            .forEach(function(foundPath){
                if (fs.lstatSync(path + "/" + foundPath).isDirectory() && !["outgoing"].includes(foundPath)){
                    var validPath = false;
                    for (var storedPathKey in groupedPaths) {
                        if (foundPath.includes(storedPathKey)){
                            groupedPaths[storedPathKey].push(path + "/" + foundPath);
                            validPath = true;
                        }
                    }
                    if (!validPath){
                        //consider removing this check and just creating a folder for this path
                        console.log("Missing path " + foundPath + " usually this means the sub paths contained a document that did not exist in base path");
                    }
                }
            });
    }
}

function createSharedFolder(options){
    var sharedFolderPath = `${options.outputPath}/shared`;
    var styleSheetPath = `${sharedFolderPath}/SkylineStyles.css`;


    return new Promise((resolve, reject) => {
        if (!fs.existsSync(options.outputPath)) fs.mkdirSync(options.outputPath);
        if (!fs.existsSync(sharedFolderPath)) fs.mkdirSync(sharedFolderPath);

        locales.forEach(function(locale){
            var localePath = `${sharedFolderPath}/${locale}`;
            if (!fs.existsSync(localePath)) fs.mkdirSync(localePath);
        })

        if(options.existingSharedPath){
            storeExistingSharedImages(options.existingSharedPath)
        }

        if(options.existingSharedPath && path.resolve(sharedFolderPath) !== path.resolve(options.existingSharedPath)){
            fs.cpSync(options.existingSharedPath, sharedFolderPath, { recursive: true });
        }
        fs.copyFileSync(`${__dirname}/skyline-resources/SkylineStyles.css`, styleSheetPath);
        fs.copyFileSync(`${__dirname}/skyline-resources/skyline.js`, `${sharedFolderPath}/skyline.js`);
        resolve();
    })
}

function storeExistingSharedImages(existingSharedPath) {
    fs.readdirSync(existingSharedPath, { recursive: true })
        .filter(
            function(file) {
                return file.endsWith(".png");
            })
        .forEach(function(imagePath){
            imagePath = imagePath.replace(/\\/g,'/')
            const fileDataBase64 = fs.readFileSync(`${existingSharedPath}/${imagePath}`, 'base64')
            var hash = crypto.createHash('md5').update(fileDataBase64).digest('hex');
            sharedImageNames[hash] = imagePath
        });
}

module.exports = {
    main: main,
}
