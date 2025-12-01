# Requirements

## Overview

This Python service connects to MS OneDrive to collect a document, reads the document, then sends the document to OpenAI to generate a summary.

## Retrieve LEFT-OFF.docx

This service will use MS Graph API to retrieve the LEFT-OFF.docx file from the OneDrive account. The file docs/reference-code/download-file.py is a reference implementation of this functionality. Excpet we want to save the LEFT-OFF.docx file to the folder found in the .env variable PATH_LEFT_OFF_SUMMARIZER.

## Parsing LEFT-OFF.docx

This document is split up into daily sections designated with a Heading 1 style in the format YYYYMMDD. Inside each daily section are two Heading 2: "LEFT-OFF" and "Accomplished Today". The document starts with the most recent entry at the top of the file and the oldest entry at the bottom of the file. This file is modified daily. So this the services will determine the date in the format YYYYMMDD from 8 days ago and copy all the text from that point to the top and save that text in a file called last-7-days-activities.md in the folder found in the .env varaible PATH_LEFT_OFF_SUMMARIZER.

## Generate Summary

This service will use OpenAI to generate a summary of the last 7 days of activities. The prompt for this service is found in a prompt.md file stored in the folder found in the .env varaible PATH_LEFT_OFF_SUMMARIZER. There is a `<< last-7-days-activities.md >>` string that should be replaced with the contents of the last-7-days-activities.md file.

## env

```
NAME_APP=LeftOffSummarizer
PATH_LEFT_OFF_SUMMARIZER=/Users/nick/Documents/_project_resources/PersonalWebsite02/left-off-summarizer
NAME_TARGET_FILE = 'LEFT-OFF.docx'
TARGET_FILE_ID = ID_TARGET_FILE
APPLICATION_ID=ID_APPLICATION
CLIENT_SECRET=CLIENT_SECRET
REFRESH_TOKEN=REFRESH_TOKEN
```
