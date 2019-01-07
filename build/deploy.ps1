# Todo: Do this multi-region
$lambdaArn = "ewsProxy";
$lambdaRegion = "eu-central-1";

Update-LMFunctionCode -FunctionName $lambdaArn -ZipFilename .\dist\dist.zip -Publish -Region $lambdaRegion
