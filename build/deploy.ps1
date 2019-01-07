# Todo: Do this multi-region
$lambdaArn = "ewsProxy";
$lambdaRegion = "eu-central-1";

yarn
yarn clean
tsc

.\build\update-version.ps1
Copy-Item -Path .\app.js -Destination .\dist
Copy-Item -Path .\version.json .\dist\
Copy-Item -Path .\package.json .\dist\
Copy-Item -Path .\yarn.lock .\dist\

cd .\dist
yarn install --production

cd ..
node .\build\zip.js
Update-LMFunctionCode -FunctionName $lambdaArn -ZipFilename .\dist\dist.zip -Publish -Region $lambdaRegion
