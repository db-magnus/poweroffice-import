# Visma Payroll to Poweroffice Go converter
This is a small python application made to be hosted in Cloud Run that takes the input from Visma Payroll and converts it to the Huldt&Lillevik format that Poweroffice Go expects.

## Install
```
gcloud auth login
gcloud config set project projectname
gcloud run deploy poweroffice-converter --source .
```
