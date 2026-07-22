# Product Dataset Analysis Using Excel

## Project Overview

This project demonstrates the use of Microsoft Excel formulas and functions to perform basic data exploration and analysis on a product dataset. The objective is to apply commonly used Excel functions for calculations, logical operations, conditional analysis, and text manipulation.

---

## Dataset

The dataset contains product-related information, including:

* Product ID
* Product Name
* Category
* Price
* Other relevant product details

---

## Tasks Performed

### 1. Basic Data Exploration

The following Excel functions were used to analyze the dataset:

#### Total Price

* **Function:** `SUM`
* **Purpose:** Calculate the total price of all products.

**Formula Example**

```excel
=SUM(C2:C101)
```

#### Number of Products

* **Function:** `COUNT`
* **Purpose:** Count the total number of products in the dataset.

**Formula Example**

```excel
=COUNT(C2:C101)
```

#### Average Product Price

* **Function:** `AVERAGE`
* **Purpose:** Calculate the average price of all products.

**Formula Example**

```excel
=AVERAGE(C2:C101)
```

---

### 2. Minimum and Maximum Price

#### Minimum Price

* **Function:** `MIN`

**Formula Example**

```excel
=MIN(C2:C101)
```

#### Maximum Price

* **Function:** `MAX`

**Formula Example**

```excel
=MAX(C2:C101)
```

---

### 3. Logical Function – IF

A new column named **Price Range** was created to classify products based on price.

**Criteria**

* Price **≥ $500** → **High Price**
* Price **< $500** → **Standard Price**

**Formula Example**

```excel
=IF(C2>=500,"High Price","Standard Price")
```

---

### 4. Conditional Functions

#### Total Price of Electronics Products

* **Function:** `SUMIF`

**Formula Example**

```excel
=SUMIF(B2:B101,"Electronics",C2:C101)
```

#### Number of Products with Price Less Than $100

* **Function:** `COUNTIF`

**Formula Example**

```excel
=COUNTIF(C2:C101,"<100")
```

---

### 5. Text Functions

Information was extracted from the **Product ID** using text functions.

#### Day

Extract the first two characters.

**Function**

```excel
=LEFT(A2,2)
```

#### Country Code

Extract the last two characters.

**Function**

```excel
=RIGHT(A2,2)
```

#### Month

Extract the 4th to 6th characters.

**Function**

```excel
=MID(A2,4,3)
```

---

## Excel Functions Used

* `SUM`
* `COUNT`
* `AVERAGE`
* `MIN`
* `MAX`
* `IF`
* `SUMIF`
* `COUNTIF`
* `LEFT`
* `RIGHT`
* `MID`

---

## Learning Outcomes

By completing this project, the following Excel skills were practiced:

* Performing basic statistical calculations
* Using logical functions for data classification
* Applying conditional aggregation functions
* Extracting text from strings using text functions
* Organizing and analyzing product data efficiently

---

## Project Structure

```
Product-Data-Analysis/
│
├── Product_Dataset.xlsx
├── README.md
└── Screenshots (Optional)
```

---

## Conclusion

This project provides hands-on experience with essential Microsoft Excel functions used in data analysis. It demonstrates how formulas can automate calculations, categorize data, summarize information, and manipulate text, making Excel an effective tool for business and data analytics tasks.


# Excel Assignment 2 – Data Exploration

## Overview

This project demonstrates **data cleaning and preprocessing techniques in Microsoft Excel** using a Product Dataset. The objective is to prepare raw data for analysis by handling missing values, correcting inconsistencies, removing duplicate records, transforming data, applying formatting, and using conditional formatting for better visualization.

---

## Dataset Information

**Dataset:** Product Dataset

### Attributes

- Product ID
- Product Name
- Brand Name
- Price ($)
- Quantity
- Category

---

## Objectives

The following data cleaning operations were performed:

- Handle missing values
- Correct inconsistent text formatting
- Fix spelling mistakes
- Remove duplicate records
- Split Product ID into separate fields
- Merge Brand Name and Product Name
- Apply number and date formatting
- Use conditional formatting for visualization

---

## Tasks Performed

### 1. Handling Missing Values

- Identified missing values in the **Price** column.
- Replaced missing prices using category-based averages.
- Filled missing categories based on product information.

---

### 2. Data Standardization

- Standardized product names using proper capitalization.
- Corrected category spelling mistakes.

Example:

| Before | After |
|---------|-------|
| laptop | Laptop |
| smartphone | Smartphone |
| Electroni | Electronics |

---

### 3. Duplicate Removal

- Identified duplicate records.
- Removed duplicate rows using Excel's **Remove Duplicates** feature.

---

### 4. Data Transformation

#### Split Product ID

Example:

```
28-JAN-US
```

into

| Manufacturing Date | Country Code |
|--------------------|--------------|
| 28-01-2024 | US |

#### Merge Columns

```
Brand Name + Product Name
```

Example:

```
Dell + Laptop
```

↓

```
Dell Laptop
```

New column:

```
Product Brand
```

---

### 5. Formatting

Applied:

- Currency formatting to Price column
- DD-MM-YYYY date formatting
- Improved table readability

---

### 6. Conditional Formatting

Applied:

- Data Bars for Price column
- Highlight rule for **Electronics** category

---

## Tools Used

- Microsoft Excel
- Excel Functions
  - PROPER()
  - CONCAT()
  - LEFT()
  - RIGHT()
  - MID()
- Find & Replace
- Remove Duplicates
- Text to Columns
- Conditional Formatting

---

## Project Structure

```
Excel-Assignment-2/
│
├── Product_Dataset.xlsx
├── Cleaned_Product_Dataset.xlsx
├── Excel_Assignment_Report.pdf
├── Screenshots/
│   ├── Missing_Values.png
│   ├── Find_Replace.png
│   ├── Duplicate_Removal.png
│   ├── Split_Columns.png
│   ├── Merge_Columns.png
│   ├── Currency_Formatting.png
│   └── Conditional_Formatting.png
└── README.md
```

---

## Learning Outcomes

This project demonstrates practical knowledge of:

- Data Cleaning
- Data Preparation
- Data Quality Management
- Data Transformation
- Data Formatting
- Spreadsheet Proficiency
- Data Visualization using Excel

---

## Skills Demonstrated

- Missing Value Handling
- Data Standardization
- Duplicate Removal
- Text Manipulation
- Data Transformation
- Number Formatting
- Date Formatting
- Conditional Formatting

---

## Screenshots

Include screenshots of:

- Missing Value Handling
- Find & Replace
- Duplicate Removal
- Split Columns
- Merge Columns
- Currency Formatting
- Conditional Formatting

---

## Conclusion

The Product Dataset was successfully cleaned and transformed using Microsoft Excel. All missing values, inconsistencies, duplicate records, formatting issues, and structural problems were resolved, making the dataset accurate, consistent, and ready for further analysis.

---

## Author

**Name:** *Your Name*

**Course:** Data Analytics

**Assignment:** Excel Assignment 2 – Data Exploration






