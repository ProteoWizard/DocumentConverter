var ArgumentParser = require("argparse").ArgumentParser;
var skyline = require("./skyline");

var parser = new ArgumentParser({
    addHelp: true
});
parser.addArgument(["--sourcePath"], {
    type: "string",
    help: "source path for docx files"
});

parser.addArgument(["--outputPath"], {
    type: "string",
    help: "output path for generated files"
});

parser.addArgument(["--existingSharedPath"], {
    type: "string",
    help: "existing tutorials path with shared resources"
});

var sourcePath = parser.parseArgs()["sourcePath"];
var outputPath = parser.parseArgs()["outputPath"];
var existingSharedPath = parser.parseArgs()["existingSharedPath"];
var options = {
    sourcePath: sourcePath,
    outputPath: outputPath,
    existingSharedPath: existingSharedPath
};
skyline.main(options);
