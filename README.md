# Docx to Html Converter

Forked from https://github.com/mwilliamson/mammoth.js

This fork was created in efforts to migrate Skyline tutorials from docx to html.
In this effort we found the need for additional features or deviations from mammoth.js.
This project can be used for any new documents that require this migration and this
documentation will cover how to do that using the same process and settings for the
original migration.

## Requirements

.NET8 - Used to convert EMF/WMF images to PNG

NodeJS - mammoth-js was written for nodejs

## Usage

To convert a directory just pass on the source and output paths:

    node .\lib\skyline-cli.js --options '
    {
        ""sourcePath"": ""C:/Users/Eduardo/Documents/Tutorials"",
        ""outputPath"": ""C:/Users/Eduardo/Documents/ConvertedTutorials""
    }'

If no output path is provided conversion will be done in place and is destructive as it rearranges files.

If you would like shared images to come from an existing shared image path then use the `existingSharedPath` option

    node .\lib\skyline-cli.js --options '
    {
        ""sourcePath"": ""C:/Users/Eduardo/Documents/Tutorials"",
        ""outputPath"": ""C:/Users/Eduardo/Documents/ConvertedTutorials"",
        ""existingSharedPath"": ""C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials/shared""
    }'

The original skyline document migration used the following options:

    node .\lib\skyline-cli.js --options '
    {
        ""sourcePath"": ""C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials"",
        ""sharedImageNameOverride"": {
            ""837cff4e0a6efc52242d8c54349ad19f"": ""en/skyline-blank-document.png"",
            ""2d6e01d456e8407ec965f49481430385"": ""zh-CHS/skyline-blank-document.png"",
            ""a40aef0d11e45f571061a78a35e2b27d"": ""ja/skyline-blank-document.png"",
            ""f5bcf8c079d513fc4c30816f2a992e41"": ""en/proteomics-interface.png"",
            ""e770486817d0cefc045395dd38ded903"": ""zh-CHS/proteomics-interface.png"",
            ""809baf27bc85792524f0fdcc4ada5c3b"": ""ja/proteomics-interface.png"",
            ""8bf47e4eb29563d8f2d3db1046375e3a"": ""en/molecule-interface.png"",
            ""a36eae3008f7cc932f6717963ac5f523"": ""zh-CHS/molecule-interface.png"",
            ""458de20a137e95913a5b228421dac36d"": ""ja/molecule-interface.png"",
            ""df6dd2f25587335b4797e7031a50e6b9"": ""molecule-icon.png"",
            ""fb4dbc95d292fb4a388cf7348d7163d1"": ""protein-icon.png"",
            ""db31e7bd65781726d0fa9313065c7555"": ""right-side-docking-icon.png""
        },
        ""subPaths"": [""Chinese"", ""Japanese""],
    }'

Where `sharedImageNameOverride` was used to identify which images we would use as a set of shared images based on their
md5 hash. It is also possible to use the `sharedImageThreshold` option to pass in an integer which creates shared images
based on a count of images it sees X amount of times.