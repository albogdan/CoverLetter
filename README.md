# Cover Letter Generator

This is a generator that takes in a CLTemplate.docx (CL = Cover Letter) file and a template.csv file and generates a cover letter with the keywords substituted in.

## Files
`CLTemplate.docx` - Your Cover Letter Template
`template.csv` - Template mapper that has all of the parameters you want to replace in your Cover Letter. Each line in the CSV corresponds to a customized Cover Letter.
`template.py` - Mapper
`README.md` - This file

## Environment Setup

### To initialize the virtual environment 'venv':
`python -m venv venv`

### To activate the virtual environment:
`.\venv\Scripts\activate`

### To deactivate the virtual environment:
`.\venv\Scripts\deactivate`

### To install requirements for the virtual environment:
`pip install -r requirements.txt`

## Instructions
Create your CLTemplate.docx with whatever you want to write in your Cover Letter. I'll leave this up to you. 
In order for this to work, all fields you want replaced must be delimited by a '{{' before the word and a '}}' after.
For example, to use the word COMPANY as a keyword, it should be written in your Cover Letter template as {{COMPANY}}
Repeat this for all keywords you want to replace

Now that your Cover Letter is done, go to your template.csv and do as follows:
1. As the first line of your template.csv, insert all the keywords you used in your CLTemplate.docx. Order doesn't matter.
2. Each line after is the values you want subsituted for each new company.
3. I recommend at least a COMPANY and a POSITION keyword, as they are how I name the files :) (you can always change that in the `template.py` file

Once you have all your values entered, you can check out the template.py file to customize:
1. Margins (Top, Bottom, Left, Right - in inches)
2. Line Spacing
3. Font Name - Note that this is kept constant throughout the document
4. Font Size - Note that this is kept constant throughout the document

Once you have adjusted your parameters to your liking you can run the program.
In the debug output is first the keywords and then each company it finds in the list

All completed Cover Letters will be found in the letters/ folder
