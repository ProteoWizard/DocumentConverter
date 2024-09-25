var ArgumentParser = require("argparse").ArgumentParser;
var skyline = require("./skyline");

var parser = new ArgumentParser({
    addHelp: true
});

parser.addArgument(["--options"], {
    type: "string",
    help: "json string options"
});

var options = JSON.parse(parser.parseArgs()["options"]);
skyline.main(options);
