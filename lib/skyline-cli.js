var ArgumentParser = require("argparse").ArgumentParser;
var main = require("./skyline");

var parser = new ArgumentParser({
    addHelp: true
});

parser.addArgument(["--options"], {
    type: "string",
    help: "json string options"
});

var options = JSON.parse(parser.parseArgs()["options"]);
main(options);
