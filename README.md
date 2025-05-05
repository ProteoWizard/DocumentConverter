# Docx to Html Converter

Forked from https://github.com/mwilliamson/mammoth.js

This fork was created in efforts to migrate Skyline tutorials from docx to html.
In this effort we found the need for additional features or deviations from mammoth.js.
This project can be used for any new documents that require this migration.

## Requirements

.NET8 - Used to convert EMF/WMF images to PNG

NodeJS - mammoth-js was written for nodejs

## Usage

Clone the repo and install node modules

    npm install

To convert a directory just input a source path to convert:

    node ./lib/skyline-cli.js --sourcePath 'C:/Users/Eduardo/Documents/TutorialSandbox'

Files will be written to an `output` folder relative to the source path. You can also supply an output path:
    
    node ./lib/skyline-cli.js --sourcePath 'C:/Users/Eduardo/Documents/TutorialSandbox' --outputPath 'C:/Users/Eduardo/Documents/TutorialSandbox/output'

If you would like shared images to come from an existing shared image path then use the `existingSharedPath` option:

    node .\lib\skyline-cli.js --sourcePath 'C:/Users/Eduardo/Documents/TutorialSandbox' --existingSharedPath 'C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials/shared'