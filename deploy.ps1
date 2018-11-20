# Todo: Do this multi-region
$lambdaArn = "arn:aws:lambda:eu-central-1:893018210320:function:ewsProxy";

yarn clean
tsc
yarn install --production --modules-folder .\dist\node_modules
.\update-version.ps1
Copy-Item -Path .\app.js -Destination .\dist
Copy-Item -Path .\version.json .\dist\
Compress-Archive -Path .\dist\* -CompressionLevel Fastest -DestinationPath .\dist\dist 
Update-LMFunctionCode -FunctionName $lambdaArn -ZipFilename .\dist\dist.zip -Publish $true