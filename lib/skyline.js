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

var imageContentTypeExtensionMap = {
    'image/png': '.png',
    'image/x-emf': '.emf',
    'image/x-wmf': '.wmf',
    'image/jpeg': '.jpeg'
}

var locales = ["en", "ja", "zh-CHS"]
var invariantFileNames = [ "invariant", "invariantDraft" ]
var htmlName = "index.html";
var regenerateImages = true;

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

var mammothStyleMap = [
    "p[style-name='Title'] => h1.document-title",
    "p[style-name='Bibliography'] => p.bibliography:fresh",
    "p[style-name='Bibliography1'] => p.bibliography:fresh",
    "p[style-name='Subtitle'] => p.subtitle",
    "r[style-name='Subtle Emphasis'] => b",
]

//stateful
var imagesToConvert = [];
var sharedImageNames = {}
var localImageNames = {}


//var debugSourcePath = "C:\\Users\\Eduardo\\Documents\\TutorialSandbox"
//var debugOutputPath = "C:\\Users\\Eduardo\\Documents\\TutorialSandbox\\output"
//var debugExistingSharedPath = "C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials/shared"
//convertFolder(debugSourcePath, debugOutputPath, debugExistingSharedPath);

//var debugConvertTutorial = "C:\\Users\\Eduardo\\Documents\\TutorialSandbox\\output\\SkylineImportingAssayLibraries\\en\\Skyline Importing Assay Libraries.docx"
//var debugConvertTutorialSharedPath = "C:\\Users\\Eduardo\\Documents\\TutorialSandbox\\output\\shared"
//convertTutorial(debugConvertTutorial, debugConvertTutorialSharedPath);

function convertFolder(sourcePath, outputPath, existingSharedPath){
    validatePathInput(sourcePath, "sourcePath")
    validatePathInput(outputPath, "outputPath")

    createSharedFolder(existingSharedPath, outputPath).then(function() {
        convertAllDocx(sourcePath, outputPath, convertDocument)
            .then(function () {
                postProcess(sourcePath);
            })
    })
}

function convertTutorial(tutorialPath, existingSharedPath){
    validatePathInput(tutorialPath, "tutorialPath")
    validatePathInput(existingSharedPath, "existingSharedPath");

    var tutorialDir = path.dirname(tutorialPath);
    var tutorialDocument = path.basename(tutorialPath);
    deleteExistingScreenshots(tutorialDir);
    storeExistingSharedImages(existingSharedPath);
    storeExistingLocalImages(tutorialDir);
    convertDocument(tutorialDir, tutorialDir, tutorialDocument)
        .then(convertImages)
        .then(function () {
            postProcess(tutorialDir);
        })
}

function validatePathInput(param, name){
    if (typeof param !== 'string' || param.trim() === '') {
        throw new Error(`Invalid input: ${name} must be a non-empty string`);
    }
}

function postProcess(sourcePath){
    var files = fs.readdirSync(sourcePath, { recursive: true });
    files.filter(
        function(file) {
            return file.endsWith(".html");
        }).forEach(function(sourceDocument) {
            var filePath = `${sourcePath}/${sourceDocument}`;
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

    return html;
}

function convertAllDocx(sourcePath, outputPath, convertFunction){
    var documentConversionPromises = [];
    documentConversionPromises.push(...convertDirectory(sourcePath, outputPath, convertFunction))

    return Promise.all(documentConversionPromises)
        .then(convertImages);
}

function convertImages(){
    //converting images at the time of document generation causes file conflict issues
    //to avoid this we convert all the images once the document generation is complete
    var imageConversionPromises = [];
    imagesToConvert.forEach(function(image){
        imageConversionPromises.push(convertToPng(image));
    })
    return Promise.all(imageConversionPromises);
}

function convertDirectory(sourcePath, outputPath, convertFunction){
    var documentConversionPromises = [];
    if (!fs.existsSync(sourcePath)) return [];
    if (!fs.existsSync(outputPath)) fs.mkdirSync(outputPath);
    var files = fs.readdirSync(sourcePath, { recursive: false });
    files.filter(
        function(file) {
            return file.endsWith(".docx") && !file.includes("~$");
        }).forEach(function(sourceDocument) {
            var documentName = path.parse(sourceDocument).name.replace(/\s+/g, '');
            var parentFolder = outputPath + "/" + documentName;
            var outputFolder = parentFolder + "/en";
            if (!fs.existsSync(parentFolder)) fs.mkdirSync(parentFolder);
            if (!fs.existsSync(outputFolder)) fs.mkdirSync(outputFolder);

            documentConversionPromises.push(convertFunction(sourcePath, outputFolder, sourceDocument));
        });
    return documentConversionPromises;
}

function convertDocument(sourcePath, outputPath, sourceDocument) {
    var documentName = path.parse(sourceDocument).name.replace(/\s+/g, '');
    var outputFileName = invariantFileNames.includes(documentName) ? `${documentName}.html` : htmlName;
    var sourceFile = sourcePath + "/" + sourceDocument;
    var outputFile = outputPath + "/" + outputFileName;
    var imageCounter = 1;

    var mammothOptions = {
        styleMap: mammothStyleMap,
        convertImage: mammoth.images.imgElement(function(image) {
            var hash;
            return image.readAsBase64String().then(function(base64Image) {
                hash = crypto.createHash('md5').update(base64Image).digest('hex');
                return image;
            }).then(function(image) {
                return image.readAsBuffer().then(function(imageBuffer) {
                    var imageLinkName;
                    if(sharedImageNames[hash]){
                        imageLinkName = "../../shared/" + sharedImageNames[hash];
                    } else if(localImageNames[hash]){
                        imageLinkName = localImageNames[hash];
                    }
                    else{
                        if (!(image.contentType in imageContentTypeExtensionMap)) {
                            throw "Unknown image type";
                        }
                        var extension = imageContentTypeExtensionMap[image.contentType];
                        var imageName = "s-" +imageCounter.toString().padStart(2, '0');
                        imageCounter = imageCounter + 1;
                        var imageFileName = imageName +extension;
                        imageLinkName = imageFileName;

                        //rename link to png as we are converting these files
                        if(extension === ".emf" || extension === ".wmf"){
                            imageLinkName = imageName + ".png"
                            if(!invariantFileNames.includes(documentName) && regenerateImages) {
                                imagesToConvert.push(outputPath + "/" +imageFileName)
                            }
                        }
                        if(!invariantFileNames.includes(documentName) && regenerateImages) {
                            fs.writeFileSync(outputPath + "/" + imageFileName, imageBuffer);
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

    return mammoth.convertToHtml({path: sourceFile}, mammothOptions)
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
            if(!(outputPath === sourcePath)){
                fs.copyFileSync(sourcePath + "/" + sourceDocument, outputPath +"/" + sourceDocument)
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

function createSharedFolder(existingSharedPath, outputPath){
    var sharedFolderPath = `${outputPath}/shared`;
    var styleSheetPath = `${sharedFolderPath}/SkylineStyles.css`;
    var scriptPath = `${sharedFolderPath}/skyline.js`;

    return new Promise((resolve, reject) => {
        if (!fs.existsSync(outputPath)) fs.mkdirSync(outputPath);
        if (!fs.existsSync(sharedFolderPath)) fs.mkdirSync(sharedFolderPath);

        locales.forEach(function(locale){
            var localePath = `${sharedFolderPath}/${locale}`;
            if (!fs.existsSync(localePath)) fs.mkdirSync(localePath);
        })

        if(existingSharedPath){
            storeExistingSharedImages(existingSharedPath)
        }

        if(existingSharedPath && path.resolve(sharedFolderPath) !== path.resolve(existingSharedPath)){
            fs.cpSync(existingSharedPath, sharedFolderPath, { recursive: true });
        } else {
            fs.copyFileSync(`${__dirname}/skyline-resources/SkylineStyles.css`, styleSheetPath);
            fs.copyFileSync(`${__dirname}/skyline-resources/skyline.js`, scriptPath);
        }
        resolve();
    })
}

function deleteExistingScreenshots(tutorialPath){
    const files = fs.readdirSync(tutorialPath);

    // Filter and delete in sync
    files
        .filter(f => /^s-\d+\.png$/.test(f))
        .forEach(f => {
            fs.unlinkSync(path.join(tutorialPath, f));
            console.log(`Deleted ${f}`);
        });
}

function storeExistingSharedImages(existingSharedPath) {
    storeExistingImagesHash(existingSharedPath, sharedImageNames);
}

function storeExistingLocalImages(tutorialPath) {
    storeExistingImagesHash(tutorialPath, localImageNames);
}

function storeExistingImagesHash(imageDir, map) {
    fs.readdirSync(imageDir, { recursive: true })
        .filter(
            function(file) {
                return file.endsWith(".png");
            })
        .forEach(function(imagePath){
            imagePath = imagePath.replace(/\\/g,'/')
            const fileDataBase64 = fs.readFileSync(`${imageDir}/${imagePath}`, 'base64')
            var hash = crypto.createHash('md5').update(fileDataBase64).digest('hex');
            map[hash] = imagePath
        });
}

module.exports = {
    convertFolder: convertFolder,
    convertTutorial: convertTutorial
}
