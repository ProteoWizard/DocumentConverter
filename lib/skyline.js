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

var options = {
    sourcePath: "C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials",
    sharedImageNameOverride: {
        "837cff4e0a6efc52242d8c54349ad19f": "en/skyline-blank-document.png",
        "2d6e01d456e8407ec965f49481430385": "zh-CHS/skyline-blank-document.png",
        "a40aef0d11e45f571061a78a35e2b27d": "ja/skyline-blank-document.png",
        "f5bcf8c079d513fc4c30816f2a992e41": "en/proteomics-interface.png",
        "e770486817d0cefc045395dd38ded903": "zh-CHS/proteomics-interface.png",
        "809baf27bc85792524f0fdcc4ada5c3b": "ja/proteomics-interface.png",
        "8bf47e4eb29563d8f2d3db1046375e3a": "en/molecule-interface.png",
        "a36eae3008f7cc932f6717963ac5f523": "zh-CHS/molecule-interface.png",
        "458de20a137e95913a5b228421dac36d": "ja/molecule-interface.png",
        "df6dd2f25587335b4797e7031a50e6b9": "molecule-icon.png",
        "fb4dbc95d292fb4a388cf7348d7163d1": "protein-icon.png",
        "db31e7bd65781726d0fa9313065c7555": "right-side-docking-icon.png"
    },
    subPaths: ["Chinese", "Japanese"],
}
//
// var invariantFolderOptions = {
//     sourcePath: "C:\\Users\\Eduardo\\Documents\\InvariantTutorials",
//     outputPath: "C:\\Users\\Eduardo\\Documents\\ConvertedInvariantTutorials",
//     existingSharedPath: "C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials/shared",
//     htmlName: "invariant.html",
// }

//main(options);
//main(invariantFolderOptions);


function main(options){
    options = setDefaultOptions(options)
    createSharedFolder(options).then(function() {
        convertAll(options, storeSharedImages).then(function() {
            convertAll(options, convertDocument).then(function () {
                rearrangeDirectoryStructure(options);
            })
        })
    })
}

function setDefaultOptions(options){
    if(!options.subPaths) options.subPaths = [];
    if(!options.htmlName) options.htmlName = "index.html";
    if(!options.sharedImageThreshold) options.sharedImageThreshold = defaultSharedImageThreshold;
    if(!options.outputPath) options.outputPath = options.sourcePath;
    return options;
}

function convertAll(options, convertFunction){
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
    var files = fs.readdirSync(currentSourcePath);
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
    var outputFileName = options.htmlName;
    var outputFolder = currentOutputPath + "/" + documentName;
    var outputFile = outputFolder + "/" + outputFileName;
    var imageCounter = 0;
    if (!fs.existsSync(outputFolder)) fs.mkdirSync(outputFolder);

    var mammothOptions = {
        styleMap: [
            "p[style-name='Title'] => h1.document-title",
            "p[style-name='Bibliography'] => p.bibliography:fresh",
            "p[style-name='Bibliography1'] => p.bibliography:fresh",
            "p[style-name='Subtitle'] => p.subtitle",
            "r[style-name='Subtle Emphasis'] => strong",
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
                        var imageName = "image-" +imageCounter++;
                        var imageFileName = imageName +extension;
                        imageLinkName = imageFileName;

                        //rename link to png as we are converting these files
                        if(extension === ".emf" || extension === ".wmf"){
                            imageLinkName = imageName + ".png"
                            imagesToConvert.push(outputFolder + "/" +imageFileName)
                        }
                        fs.writeFileSync(outputFolder + "/" +imageFileName, imageBuffer);
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
            fs.cpSync(options.existingSharedPath, sharedFolderPath, { recursive: true });
            storeExistingSharedImages(options.existingSharedPath)
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

module.exports = main;
