# Simple_MCQ_checker

## Introduction
This simple tool has been designed to help you grade loads of Multiple choice question (MCQ) assessments with a single click. 
Saves you time so that you can waste more of it on Netflix :p   
**Sailent Features include:**
  * Virtually no limit on the number of questions that can be assessed in a single assesment sheet.
  * Virtually no limit on the number of assessments that can be checked.
  * Generates a consolidated result excel file.
  * Fairly commented code.
  * Easy to modify according to your own needs. (see LICENSE for terms)
  * Free to use. (see LICENSE for terms)

## Installing Dependencies

The tool runs on Python 3. To save you the hassel of installing the dependencies it is recommended that you
use the online [PythonAnywhere online IDE](https://www.pythonanywhere.com). Just create an account and you are good to go.

The tool assumes that the assesment is done using the provided "Assessment_template.xlsx" for questionnaire.  
You will need to modify the **answer_checker.py** script if you intend to use your custom template.

If you are uncomfortable uploading your files online, you can install python locally and following dependencies:
using pip from command line.
  * **openpyxl** : pip install openpyxl
  * **glob** : pip install glob2
  
## Setting up directories

If you are running this locally or even on PythonAnywhere, just do the following:
  * Create a directory say "Answer_checker".
  * Copy the following to this directory:
    * **answer_checker.py**
    * All the sheets to be assessed to **the same directory**.
    * **Assess_answer.xlsx** containing the correct answers to questions to **the same directory**.  
      All the answer should be edited in column 'A' of Assess_answer.xlsx and can be one of following: 
      * 'a.','b.','c.','d.'   
      **__Don't forget the dots__** (feel free to modify the code and questionnaire template if you need other/more options)  
      The answer for question 1 should be fed in cell A1 of excel, question 2 cell A2,question 3 cell A3
      and so on.
  * Make sure **no other** .xlsx file is present in the folder. The tool can crash otherwise.

## Running the tool

Just run the **answer_checker.py** and you are done. The tool generates a **Consolidated_result.xlsx** that contains 
the final score of each candidate.

### That's all you need to do. Feel free to suggest improvements ###
