# Data Quality Report

## 1. Project Overview

This project demonstrates a customer data entry and data quality validation workflow using Microsoft Excel.

The dataset contains 10 customer records with information including Record ID, Full Name, Age, Gender, Phone Number, Email, State, and Registration Date.

The objective was to review the dataset for common data-entry and data-quality issues, validate the information, and produce a clean, organized dataset suitable for further use.

## 2. Objectives

The main objectives of the project were to:

- Enter and organize customer information accurately.
- Identify missing or incomplete records.
- Check for duplicate customer records.
- Validate phone number formats.
- Check email addresses for basic formatting requirements.
- Validate customer age ranges.
- Verify location information using a reference table.
- Validate registration dates.
- Identify dates that occur in the future.
- Produce a clean and structured final dataset.
- Summarize the results of the data quality checks.

## 3. Dataset Description

The dataset contains 10 customer records.

Each record contains the following fields:

| Field | Description |
|---|---|
| Record ID | Unique identifier assigned to each customer record |
| Full Name | Customer's full name |
| Age | Customer's age |
| Gender | Customer's gender |
| Phone Number | Customer's contact phone number |
| Email | Customer's email address |
| State | Customer's state or location |
| Registration Date | Date the customer was registered |

The original dataset was preserved in its raw form before validation and cleaning. A separate working copy was used for data-quality checks, while the final validated data was saved as a cleaned dataset.

## 4. Data Entry and Validation Process

The customer records were entered and organized in Microsoft Excel. A separate working copy of the dataset was used to perform validation and quality checks without altering the original raw data.

The following validation procedures were performed:

### 4.1 Name Validation

Customer names were reviewed for:

- Missing names
- Duplicate names
- Inconsistent formatting
- Unnecessary spaces

The TRIM() function was used to check for unnecessary spaces in names.

### 4.2 Phone Number Validation

Phone numbers were checked for:

- Missing values
- Correct length
- Correct starting format
- Duplicate phone numbers

The LEN() function was used to check the number of characters in each phone number.

The LEFT() function was used to verify that the phone numbers started with 0.

### 4.3 Email Validation

Email addresses were checked for:

- Missing values
- Basic email formatting
- Duplicate email addresses

The SEARCH() function was used to look for the @ and . characters.

ISNUMBER() was used to determine whether the SEARCH() function successfully found those characters, while AND() ensured that both conditions were satisfied.

### 4.4 Age Validation

Customer ages were checked against an acceptable range of 18 to 60 years.

The IF() and AND() functions were used to classify each age as either acceptable or invalid.

### 4.5 Location Validation

Customer locations were validated using a separate reference table.

An Excel Table named Table1 was created on the Reference data sheet. The table contained the location and corresponding location type.

XLOOKUP() was then used to compare the location in the customer dataset against the reference table.

This allowed locations such as Abuja to be correctly classified as FCT rather than being incorrectly treated as a state.

### 4.6 Registration Date Validation

Registration dates were checked to determine whether Excel recognized them as valid date values.

The ISNUMBER() function was used because Excel stores valid dates internally as numeric serial values.

The TODAY() function was also used to identify registration dates occurring in the future.

### 4.7 Duplicate Record Validation

The Record ID, phone number, and email columns were checked for duplicates using COUNTIF().

Records were classified as either Unique or Duplicate.

### 4.8 Missing Data Validation

The COUNTA() function was used to check whether all required fields in each customer record contained data.

Records containing all required fields were classified as Complete.

## 5. Data Quality Results

A total of 10 customer records were reviewed during the data-quality validation process.

| Data Quality Check | Result |
|---|---:|
| Total Records | 10 |
| Complete Records | 10 |
| Duplicate Record IDs | 0 |
| Duplicate Phone Numbers | 0 |
| Duplicate Email Addresses | 0 |
| Invalid Email Records | 0 |
| Invalid Age Records | 0 |
| Unknown Locations | 0 |
| Invalid Registration Dates | 0 |
| Future Registration Dates | 0 |

All 10 customer records passed the validation checks performed during the project.

No duplicate Record IDs, phone numbers, or email addresses were identified. No missing required fields, invalid email formats, invalid ages, unknown locations, invalid dates, or future registration dates were identified.

## 6. Cleaning and Formatting

After the validation process, a separate cleaned dataset was created.

The final dataset was organized into the following fields:

- Record ID
- Full Name
- Age
- Gender
- Phone Number
- Email
- State
- Registration Date

The cleaned dataset was formatted as an Excel Table to improve readability, filtering, sorting, and data management.

Phone numbers were preserved in a format that retained the leading zero, while registration dates were formatted as recognizable Excel dates.

The original raw dataset was preserved separately from the cleaned dataset.

## 7. Tools Used

- Microsoft Excel
- Excel Tables
- Conditional Formatting
- XLOOKUP
- COUNTIF
- COUNTA
- IF
- AND
- ISNUMBER
- SEARCH
- LEN
- LEFT
- TRIM
- TODAY

## 8. Final Outcome

The customer dataset was successfully reviewed, validated, and organized using Microsoft Excel.

All 10 records passed the data-quality checks performed. A separate cleaned dataset was produced while preserving the original raw data.

The completed project demonstrates practical skills in data entry, data validation, data cleaning, Excel formulas, reference-table lookups, duplicate detection, and documentation.




