# Change sensitivity label in folder
## To get the files to insert extract  xlsx file with label:

1. run  for any xlsx file with desired label ```unzip myfile.xlsx -d myfile```
2. copy content of ```docProps/custom.xml``` into the code
3. Do this for each label you want to use

## If you want to use python: ```sensitivity_xlsx.py```

1. Change ```xml_content_``` variables
2. Change ```file_to_insert```
3. Install packages using conda from environment.yaml
4. Run the code

## If you want to use powershell: ```sensitivity_xlsx.ps```

1. Change ```$xmlContent``` variables
2. Change ```$fileToInsert```
3. Install powershell
4. Run the code