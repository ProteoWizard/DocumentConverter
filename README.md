# Docx to Html Converter

Forked from https://github.com/mwilliamson/mammoth.js

This fork was created in efforts to migrate Skyline tutorials from docx to html.
In this effort we found the need for additional features or deviations from mammoth.js.
This project can be used for any new or existing documents that require this migration.

## Requirements

.NET8 - Used to convert EMF/WMF images to PNG

NodeJS - mammoth-js was written for nodejs

## Usage

Clone the repo and install node modules

    npm install

To convert a new tutorial or reconvert an existing tutorial you can use the `convert-tutorial` command to target your tutorial:

    node .\lib\skyline-cli.js convert-tutorial --source-path 'C:\Users\Eduardo\repos\pwiz\pwiz_tools\Skyline\Documentation\Tutorials\ImportingAssayLibraries\en'

This directory corresponds to the [pwiz repo](https://github.com/ProteoWizard/pwiz/tree/master/pwiz_tools/Skyline/Documentation/Tutorials/DIA/en).
The command will regenerate any images it considers screenshots(s-##.png). The command will
detect if any of the images found in the docx are in the tutorial folder or in the shared
folder. This means you can re-run this command with images renamed to maintain them as non
screenshots. The `source-path` can point to the docx file or the directory containing it,
but it must be in the Skyline tutorials directory. 

It is possible to still convert a folder containing multiple docx files using the `convert-folder` command.
Output will be written to `output-path` or default to `{source-path}/output`. `existing-shared-path` provides
a way to use an existing shared path for images/resources in the new output. 

    node .\lib\skyline-cli.js convert-folder --source-path 'C:/Users/Eduardo/Documents/TutorialSandbox' --existing-shared-path 'C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials/shared'