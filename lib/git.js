var fs = require('fs');
var exec = require('child_process');

var baseGitFolder = 'C:/Users/Eduardo/repos/pwiz';
var baseTutorialPath = "C:/Users/Eduardo/repos/pwiz/pwiz_tools/Skyline/Documentation/Tutorials";

var commits = []
commit();

function commit(){

    // commitPath("docx", "Remove Old tutorials")
    // commitPath("Japanese", "Remove Japanese tutorials")
    // commitPath("Chinese", "Remove Chinese Tutorials")
    // populate base map

    commitPath(`pwiz_tools/Skyline/Documentation/Tutorials/shared`, `Add shared files`);

    var paths = [];
    fs.readdirSync(baseTutorialPath)
        .forEach(function(file){
            if (fs.lstatSync(baseTutorialPath + "/" + file).isDirectory() && !["Chinese", "Japanese", "shared"].includes(file)){
                paths.push(file);
            }
        });
    paths.sort(function (a, b) {
        return b.length - a.length;
    }).forEach(function (path) {
        commitPath(`pwiz_tools/Skyline/Documentation/Tutorials/${path}`, `Convert ${path} documents`);
        console.log(path)
    })

    commitPath(`*docx*`, `Remove old documents`);

    commits.forEach(function(commit){
        console.log(commit);
    })
}

function commitPath(path, message){

    var addPath = ["add", path];
    exec.execFileSync("git", addPath, {cwd: baseGitFolder, maxBuffer: 50000 * 1024 * 1024});

    var commitArgs = ["commit", "-m", message];
    exec.execFileSync("git", commitArgs, {cwd: baseGitFolder, maxBuffer: 50000 * 1024 * 1024});

    var getCommitIdArgs = ["rev-parse", `HEAD`];
    var commitIdBuffer = exec.execFileSync("git", getCommitIdArgs, {cwd: baseGitFolder});
    commits.push(`${path} ${commitIdBuffer +''}`);
}
