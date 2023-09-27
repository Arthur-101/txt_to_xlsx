# Data Filter and Analysis Project Documentation


# ->  Table of Contents:

1.  Introduction
2.  Project Overview
3.  Installation
4.  Usage
5.  File Structure
6.  Libraries Used
7.  Project Workflow
8.  Data Processing
9.  Data Analysis
10. Output
11. Conclusion
12. Contributors


#   1. INTRODUCTIONThis documentation provides an overview of the Filter_X and Filter_XII Python scripts, which are part of a project aimed at processing and analyzing text data from student results and saving the results in an Excel file. These scripts are designed to extract, clean, and analyze data for classes X and XII, respectively.


#    2. PROJECT OVERVIEW
The project consists of two Python scripts, `Filter_X.py` and `Filter_XII.py`, each designed to filter and process data for Class X and Class XII students, respectively. The data extraction and analysis tasks are structured into classes within these scripts, making it easy to use and maintain.


#     3. REQUIREMENTS
To run the Data Filter and Analysis project, the following requirements must be met:

-Python 3.x installed on the system.
-Required Python modules: `re`, `pandas`, `openpyxl`
-Required extentions: `Jupyter Notebook`


#      4. INSTALLATION
To install the Data Filter and Analysis project, follow these steps:

1. Ensure Python 3.10 is installed on your system.

2. Download the project files and save them to a directory of your choice.

3. Open a terminal or command prompt and navigate to the project directory.

4. Install the required Python modules by running the following command:
  pip install regex pandas openpyxl

6. Once the dependencies are installed, open the files in jupyter notebook.

7. After uploading the files, you can run `Main.ipynb` file.


#      5. USAGE
The main functionality of the project is accessed through the `Main.ipynb` Jupyter notebook. To use the project, follow these steps:

1. Import the necessary Python scripts:
    from Filter_X import DataProcessor_X
    from Filter_XII import DataProcessor_XII

2. Create instances of the `data_processor_X` and `data_processor_XII` classes, passing the path to the text files containing student data as arguments:
    data_processor_X = DataProcessor_X(input_file_X)
    data_processor_XII = DataProcessor_XII(input_file_XII)

3. Use the methods provided by these classes to filter and analyze the data.

4. Save the results in an Excel file :
    data_processor_X.save_data_to_excel(output_file_X)
    data_processor_X.save_analysis_to_excel(output_file_X)
    data_processor_XII.save_data_to_excel(output_file_XII)
    data_processor_XII.save_analysis_to_excel(output_file_XII)


#       6.  File Structure
The project has the following file structure:

- `Main.ipynb`  is the Jupyter Notebook file for calling and analyzing the scripts.
- `Filter_X.py` is the main scripts for data processing for class X.
- `Filter_XII.py` is the main scripts for data processing for class XII.
- input txt file for X  contains the input text file of class X.
- input txt file for XII contains the input text file of class XII.


#       7. DEPENDENCIES
This project depends on the following Python libraries:

`re`: For regular expression-based text data extraction.
`pandas`: For data manipulation and dataframe creation.
`openpyxl`: For working with Excel files.
`openpyxl.styles`: For styling Excel sheets.
`openpyxl.utils.dataframe`: For converting dataframes to Excel sheets.


#        8. Project Workflow
The project workflow can be summarized as follows:

1. Data extraction: Extract relevant information from the text files using regular expressions and convert it into a Pandas DataFrame.

2. Data processing: Perform data cleaning and processing, including calculating grades and percentages.

3. Data analysis: Analyze the data to generate various statistics, such as student percentage counts, subject-wise percentage counts, and highest marks.

4. Save results: Save the analyzed data in separate sheets of an Excel file.

#         9. Data Processing
Data processing involves several steps:

1. Data cleaning: Remove any inconsistencies or errors in the extracted data.

2. DataFrame : Creating DataFrame based on the cleaned data.

3. Calculations: Compute obtained data for various analysis.


#          10. Data Analysis
The project performs the following data analysis tasks:

1. percentage_counts_df: Counts the number of students falling into specific percentage ranges (e.g., Above 90%, 80%-89%, etc.).

2. subject_percentage_count_df: Counts the number of students within each percentage range for each subject.

3. highest_marks_df: Identifies students with the highest marks in each subject.


#           11. Output
The project generates an Excel file containing the following sheets:

- `Result_Analysis_X`: Contains the results of Class X students.
- `Result_Analysis_XII`: Contains the results of Class XII students.


#           12. CONTRIBUTORS
The Data Filter and Analysis project was developed by [Saurav] as a school project.


#           13. CONCLUSION
This project provides a structured and efficient way to filter and analyze student data from text files. By following the steps outlined in this documentation, users can extract valuable insights from the data and store them in an organized Excel format.
