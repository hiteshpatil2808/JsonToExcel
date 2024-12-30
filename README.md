# JSON to Excel Conversion Project

## Overview
This project, developed during my internship at iConcile Technologies Pvt. Ltd., is a Spring Boot application designed to automate the conversion of dynamic JSON data into formatted Excel spreadsheets. The application ensures the resulting Excel sheets are well-organized, easy to read, and adhere to specified formatting requirements.

## Key Features
- **Dynamic Data Handling**: Supports JSON data with varying structures, providing flexibility in data processing.
- **Custom Formatting**:
  - **Row and Column Freezing**: Enhances navigation by freezing specific rows and columns.
  - **Conditional Background Colors**: Highlights important data with color-coding based on specified conditions.
  - **Currency Formatting**: Formats specified columns as currency for consistent financial data presentation.
  - **Column Removal**: Excludes unnecessary columns to keep the Excel sheet focused and clean.
  - **Custom Labeling**: Renames columns for better clarity and understanding.
  - **Date and Number Formatting**: Applies appropriate date and number formats for consistency.

## Technologies Used
- **Backend**: Java, Spring Boot (MVC), Hibernate
- **Tools**: Eclipse, Spring Tool Suite 4, Postman
- **Libraries**: Apache POI for Excel operations

## How It Works
1. **API Development**:
   - RESTful endpoints are created to accept JSON data and initiate the conversion process.
   - Comprehensive error handling ensures robust operation.

2. **Data Processing**:
   - Incoming JSON data is parsed and validated to ensure compatibility with the conversion logic.
   - Dynamic formatting rules are applied to prepare the data for Excel generation.

3. **Excel Generation**:
   - The Apache POI library is used to create and format Excel files.
   - Logic for custom formatting, including conditional styling and currency formats, is implemented.

## Usage
1. **Clone the Repository**:
   ```sh
   git clone https://github.com/hiteshpatil2808/JsonToExcel.git
   cd JsonToExcel
