var fs = require('fs');

var images = {
    "AbsoluteQuant" : {
        "image-0" : "s-experimental-overview"
    },
    "AuditLog" : {
        "image-26" : "s-panorama-zip-view",
        "image-27" : "s-panorama-audit-log",
        "image-28" : "s-create-view",
        "image-29" : "s-customized-view"
    },
    "CustomReports" : {
        "image-5" : "copy-button"
    },
    "DDASearch" : {},
    "DIA" : {
        "image-2" : "s-precursor-frequency-graph",
        "image-3" : "s-isolation-window-graphs",
        "image-8" : "s-isolation-list-csv",
        "image-26" : "vertical-split-cursor-icon",
        "image-29" : "magnifying-glass-icon",
        "image-31" : "arrow-buttons",
        "image-34" : "right-arrow-button"
    },
    "DIA-PASEF" : {
        "image-0" : "s-pasef-label-free-proteome-quantification"
    },
    "DIA-QE" : {
        "image-0" : "s-qe-label-free-proteome-quantification"
    },
    "DIA-TTOF" : {
        "image-0" : "s-ttof-label-free-proteome-quantification"
    },
    "DIA-Umpire-TTOF" : {
        "image-0" : "s-ttof-label-free-proteome-quantification"
    },
    "ExistingQuant" : {
        "image-6" : "peptide-icon",
        "image-7" : "peptide-with-spectrum-lines-icon",
        "image-12" : "pushpin-icon",
        "image-13" : "s-precursor-total-area-ratio-formula",
        "image-15" : "s-transition-list-spreadsheet"
    },
    "GroupedStudies" : {
        "image-0" : "s-detection-and-differentiation-phases",
        "image-2" : "peptide-with-spectrum-lines-icon",
        "image-3" : "peptide-without-spectrum-lines-icon",
        "image-19" : "red-x-icon",
        "image-45" : "protein-icon-2",
        "image-84" : "red-x-icon",
        "image-86" : "red-x-icon",
        "image-104" : "red-x-icon",
    },
    "HiResMetabolomics" : {
        "image-0" : "s-transition-list-spreadsheet",
    },
    "ImportingAssayLibraries" : {
        "image-0" : "s-experiment-assay-library",
    },
    "ImportingIntegrationBoundaries" : {
        "image-0" : "s-peak-boundary-csv",
    },
    "IMSFiltering" : {
        "image-12" : "show-2d-spectrum-button",
        "image-14" : "magnifying-glass-icon"
    },
    "iRT" : {
        "image-3" : "calculator-button",
        "image-6" : "horizontal-split-cursor-icon",
        "image-22" : "s-window-size-formula"
    },
    "LibraryExplorer" : {},
    "MethodEdit" : {
        "image-23" : "s-method-edit-file-explorer",
        "image-24" : "s-transition-list-spreadsheet"
    },
    "MethodRefine" : {
        "image-0" : "s-targeted-method-refinement-cycle",
        "image-21" : "s-transition-list-spreadsheet"
    },
    "MS1Filtering" : {
        "image-2" : "s-ms1-filtering-folder",
        "image-7" : "s-relative-abundance-graph",
        "image-8" : "s-chromatogram",
    },
    "OptimizeCE" : {},
    "PeakPicking" : {
        "image-9" : "blue-binocular-button",
        "image-12" : "peptide-with-spectrum-lines-and-clock-icon",
        "image-29" : "blue-binocular-button",
    },
    "PRM" : {
        "image-0" : "s-triple-quadrupole-mass-spectrometer",
        "image-1" : "s-single-intensity-measurements",
        "image-2" : "s-full-spectrum-chromatogram",
        "image-3" : "s-intensity-retention-time-graph",
        "image-11" : "s-ltq-instrument-setup-software",
        "image-23" : "target-peptide-with-spectrum-icon"
    },
    "PRMOrbitrap" : {},
    "PRMOrbitrap-PRBB" : {
        "image-18" : "black-binocular-button",
        "image-19" : "s-precursor-list-spreadsheet",
        "image-24" : "red-dot-icon"
    },
    "SmallMolecule" : {},
    "SmallMoleculeIMSLibraries" : {
        "image-10" : "bottom-side-docking-icon",
        "image-11" : "right-side-docking-icon",
        "image-12" : "left-side-docking-icon",
        "image-16" : "show-2d-spectrum-button",
    },
    "SmallMoleculeMethodDevCEOpt" : {
        "image-0" : "s-published-transition-list"
    },
    "SmallMoleculeQuantification" : {
        "image-0" : "s-96-well-plate-layout",
        "image-1" : "s-injection-order",
        "image-14" : "vertical-split-cursor-icon",
        "image-20" : "black-binocular-button"
    }
}

var tutorialsPath = "C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials";
var sharedPath = "C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials/shared";
renameImages();


function renameImages(){
    Object.keys(images).forEach(function(tutorialName){
        processHtml(tutorialName, "en", "index.html");

        if(fs.existsSync(`${tutorialsPath}/${tutorialName}/ja`)){
            processHtml(tutorialName, "ja", "index.html");
            processHtml(tutorialName, "ja", "invariant.html");
        }

        if(fs.existsSync(`${tutorialsPath}/${tutorialName}/zh-CHS`)){
            processHtml(tutorialName, "zh-CHS", "index.html");
            processHtml(tutorialName, "zh-CHS", "invariant.html");
        }
    });
}

function processHtml(tutorialName, locale, fileName){
    console.log(`Processing ${tutorialName} ${locale} ${fileName}`);
    var currentTutorialPath = `${tutorialsPath}/${tutorialName}`
    var html = fs.readFileSync(`${currentTutorialPath}/${locale}/${fileName}`, "utf8");

    Object.entries(images[tutorialName]).forEach(([imageName, imageReplacementName]) => {
        if(!html.includes(`${imageName}.png`)){
            throw `Could not find image ${imageName} in tutorial ${tutorialName}`
        }

        var isSharedImage = !imageReplacementName.startsWith("s-");
        var originalImagePath = `${currentTutorialPath}/${locale}/${imageName}.png`;
        var destinationImagePath = isSharedImage ? `${sharedPath}/${imageReplacementName}.png`: `${currentTutorialPath}/${locale}/${imageReplacementName}.png`
        var originalImageText = `${imageName}.png`
        var destinationImageText = isSharedImage ? `../../shared/${imageReplacementName}.png` : `${imageReplacementName}.png`

        html = html.replace(originalImageText, destinationImageText);

        if(fileName !== "index.html"){
            //only change file structure for index.html
        }
        else if(!fs.existsSync(destinationImagePath)){
            fs.renameSync(originalImagePath, destinationImagePath);
        }else {
            fs.rmSync(originalImagePath)
        }
    });

    var oldImagesFileNames = html.match(/image-[^'"]+/g) || []
    var newImageCounter = 1;
    var newImageFileNames = {} //used for images that are reused
    oldImagesFileNames.forEach(function (oldImageFileName) {

        var newImageFileName;
        if(newImageFileNames[oldImageFileName]){
            newImageFileName = newImageFileNames[oldImageFileName]
        } else {
            newImageFileName = `s-${newImageCounter++}.png`
            newImageFileNames[oldImageFileName] = newImageFileName
            if(fileName === "index.html"){
                //only change file structure for index.html
                fs.renameSync(`${currentTutorialPath}/${locale}/${oldImageFileName}`, `${currentTutorialPath}/${locale}/${newImageFileName}`);
            }
        }

        html = html.replace(oldImageFileName, newImageFileName);
    })
    fs.writeFileSync(`${currentTutorialPath}/${locale}/${fileName}`, html, "utf8");

    var filesInTutorialFolder = fs.readdirSync(`${currentTutorialPath}/${locale}`);
    filesInTutorialFolder.filter(
        function(foundFile) {
            return foundFile.startsWith("image-");
        }).forEach(function(foundFile) {
        console.log(`Found unused image file ${foundFile}`)
    })
}
