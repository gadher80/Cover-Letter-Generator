# Cover Letter Generator

This repository contains an application that generates a cover letter based on user inputs. 
The inputs include the name of the company, the name of the position, the address of the company, the person to be greeted, and the domain for which the cover letter needs to be created (Data Science, Data Analytics, Machine learning, or Software Development). 
The output can be either a PDF of the whole job application combined with the cover letter and CV or a CV and DOCX file of the cover letter.

## Usage
To use the application, download or clone the repository to your local machine. Navigate to the repository directory in your terminal or command prompt and install the required libraries by running the following command:


## Run

```pip install -r requirements.txt```

1. After installing the required libraries run the coverUI3.py file to start the application.

2. Once the application is running, fill out the required fields and select the checkbox if you want to use a predefined CV. Then, click the "Generate Cover Letter" button to generate the cover letter. The generated cover letter will be saved in the same directory as a DOCX file, and a PDF of the whole job application (including cover letter and CV) will be saved in the same directory.

## Files

1. coverUI3.py: The main Python script that contains the application code.
2. CoverLetter.docx: The default cover letter template file.
3. default_CV.pdf: The default CV file.
4. requirements.txt: The file containing the required libraries for the application.

## Conclusion

The Cover Letter Generator is a user-friendly PyQt5 UI application that generates a cover letter based on user inputs. The application allows the user to generate either a PDF of the whole job application or a DOCX file of the cover letter. The predefined CV feature adds convenience for users who do not want to manually attach their CV to the application.
