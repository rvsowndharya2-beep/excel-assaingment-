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

