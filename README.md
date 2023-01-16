# ResearchAMSIProvider
Simple AMSI Provider that writes input to tmp files

## Usage
1. Register the COM object
```
regsvr32.exe ResearchAMSIProvider.dll
```

That's all! Now if you execute powershell/macros/vbscript the content of that script will be written to corresponding file in your %TMP% directory.

To Uninstall, run
```
regsvr32.exe /u ResearchAMSIProvider.dll
```
