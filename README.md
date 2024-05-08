# SAP-GUI-Python-Scripting

Author : Freek Klabbers\
Date : 8-5-2024\

## Introduction
This (poorly constructed) repo is meant as a storage and a showcase of how Python can be used to automate SAP ERP. (only applied to SAP Plant Maintenance (PM) but still).\
This is not in any way ready to go software, but more a generic example of a few things I have used the feature for in the past.

## Structure
The SAP_Automation_Class is the main feature of this repo, it contains a Parent class to all other programs with a generic setup to get SAP up and running. \
With its method you should be able to get like 70% of the way there. Then Child classes can be used to overwrite a few of the methods.

Every other Class is an instance of a use-case I have used in the past.\
A lot of these come down to "I need to add a lot of data to a lot of different equipments but i dont wanna do it by hand".\
For this I just made a loop that runs over an excel document and extracts the data and inputs it into SAP.\

Login will be different for every user so check this out for yourself. I used the Single Sign On (SSO) feature, which means i dont have to store a password in code, or have to work with secrets.\
Highly recomnd to follow this approach if possible.

## Way of working
What i did is, whenever there was an opperation I needed to follow, I recorded it with the record feature in SAP, and then took the VBA code and just pasted into Python, from there you just need to add some parentheses to a few methods and it works. This way you dont need to know the name of every button and controll.
