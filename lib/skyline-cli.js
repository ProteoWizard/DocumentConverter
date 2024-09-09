var ArgumentParser = require("argparse").ArgumentParser;
var main = require("./skyline");

var parser = new ArgumentParser({
    addHelp: true
});

parser.addArgument(["--options"], {
    type: "string",
    help: "json string options"
});

for (let i = 0; i < process.argv.length; ++i) {
    console.log(
        `index ${i}
        argument ->
        ${process.argv[i]}
        `
    );
}

console.log(parser.parseArgs()["options"]);
var options = JSON.parse(parser.parseArgs()["options"]);
main(options);
