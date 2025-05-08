var ArgumentParser = require("argparse").ArgumentParser;
var skyline = require("./skyline");
var fs = require('fs');
var  path = require('path');

var parser = new ArgumentParser({
    addHelp: true
});


var subCommandParser = parser.addSubparsers({
    title: 'subcommands',
    dest: 'command',
    required: true
});


var convertTutorial = subCommandParser.addParser('convert-tutorial', {addHelp: true});

convertTutorial.addArgument(['--source-path'], {
    help: 'Tutorial path to convert',
    required: true
});


var convertFolder = subCommandParser.addParser('convert-folder', {addHelp: true});

convertFolder.addArgument(["--source-path"], {
    type: "string",
    help: "Source path for docx files",
    required: true
});

convertFolder.addArgument(['--existing-shared-path'], {
    type: "string",
    help: 'Path to existing shared folder, will not use one if not provided',
});

convertFolder.addArgument(["--output-path"], {
    type: "string",
    help: "Output path for generated files, will default to output folder relative to source"
});



var args = parser.parseArgs();
var command = args['command'];

switch (command) {
    case 'convert-folder':
        var sourcePath = args['source_path'];
        var outputPath = args['output_path'];
        if(!outputPath){
            outputPath = sourcePath + "/output"
        }
        var existingSharedPath = args['existing_shared_path'];
        skyline.convertFolder(sourcePath, outputPath, existingSharedPath);
        break;
    case 'convert-tutorial':
        var sourcePath = args['source_path'];
        var tutorialPath = findTutorialPath(sourcePath);
        console.log(`Using ${tutorialPath} as input file`)
        var existingSharedPath = findExistingSharedPath(sourcePath);
        console.log(`Detected the shared directory ${existingSharedPath}`)
        skyline.convertTutorial(tutorialPath, existingSharedPath);
        break;
    default:
        throw new Error(`Unsupported command: ${command}`);
}

function findExistingSharedPath(sourcePath){
    let currentDir = path.resolve(sourcePath);

    var maxDepth = 5;
    for (let i = 0; i < maxDepth; i++) {
        const sharedPath = path.join(currentDir, 'shared');

        if (fs.existsSync(sharedPath) && fs.statSync(sharedPath).isDirectory()) {
            return sharedPath;
        }

        const parentDir = path.dirname(currentDir);
        if (parentDir === currentDir) {
            break; // Reached the root
        }

        currentDir = parentDir;
    }

    throw new Error(`Could not detect shared path when starting from path ${sourcePath}`);
}

function findTutorialPath(sourcePath){
    if(sourcePath.toLowerCase().endsWith(".docx")){
        console.log(`Input path is a docx file ${sourcePath}.`)

        if(fs.existsSync(sourcePath)){
            console.log("Validated file exists.")
            return sourcePath;
        } else {
            throw new Error(`Input file does not exist`);
        }
    } else if(fs.statSync(sourcePath).isDirectory()){
        var docxFiles = fs.readdirSync(sourcePath)
            .filter(file => !file.startsWith("~$"))
            .filter(file => path.extname(file).toLowerCase() === '.docx');

        console.log(docxFiles)
        if (docxFiles.length === 0) {
            throw new Error(`No .docx file found in directory ${sourcePath}`);
        } else if (docxFiles.length > 1) {
            throw new Error(`More than one .docx file found in directory. Found: ${docxFiles}.`);
        }
        return path.join(sourcePath, docxFiles[0]);
    }

    throw new Error(`source-path ${sourcePath} not a valid file or directory`);
}