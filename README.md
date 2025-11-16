## Invoice Generator (GitHub Version)

This repository contains a Python-based invoice generator with an interactive UI, built using ipywidgets. The GitHub version is functionally the same as the internal version, but uses anonymized data for confidentiality. Most of the code is organized into functions to support dynamic widgets, allowing interactive invoice creation, filtering, and formatting. This is based on Jupyter, hence the `ipywidgets` usage. This is the reason, why there is .py file for now.

---

## Output demo

![Invoice generation demonstration](images/Demo.gif)

---


## Features:

#### Data Preparation  
- Loads a multi-sheet Excel file which is used to extract and clean relevant columns. converts sheet data into product group dataframes and renames columns into English for consistency.  

#### Product Filtering System  
- Regex-based filtering for each product category which allows dynamic filter assignment via dictionaries. 
- Multi-layer filter logic (category → subgroup → custom filter)  
- “Common items” filtering mode, which allows to narrow down the product selection to specific amount by groups.

#### Invoice Generator Logic  
- The invoice generator uses a probability-driven algorithm to decide which products appear, their quantities, total sum distribution and how the final invoice meets user-defined price bounds. This means that this algorithm uses upper and lower price bounds, different generation modes, like normal or 'common' and minimum quantity with spread factor to guarantee that the invoice total meets the user-defined lower bound. This section describes the behavior based on the `Product_price_generator` implementation. 

#### Fully Interactive UI (ipywidgets)  
- Fully working UI, that contains category selection checkboxes, sub-filter selection UI that updates live auto-adjusting sliders based on selected filters, dynamic visibility of UI elements based on observers and client selection & custom client entry if needed. 

#### Excel Invoice Output  
- Generates full Excel invoice using `openpyxl`. Generation includes automatic row resizing based on item count, custom formatting (fonts, alignment, margins, borders). Generated invoices also support the conversion of the invoice's total into words and error-safe file naming and export system.

---

## Code Structure

1. **Spreadsheet Preparation** - Separate sheets and creates cleaned dataframes.

2. **Dataframe Grouping** - Product groups created using regex and indexed segmentation.

3. **Generator Logic** - Quantity and price generation using probability models.

4. **Filter Creation** - Category and sub-filter creation utilizing checkboxes using regex patterns.

5. **User Interface** - Ipywidgets-based UI with dynamic observers.

6. **Custom Filtering System** - Builds filter masks to update product dataframe.

7. **Slider Logic** - Sliders adjust automatically based on selected filters.

8. **Invoice Settings** - Upper/lower price bounds and target goal calculation.

9. **Client Data Handling** - Client information handling and storage

10. **Excel Formatting** - Formats invoice using openpyxl and outputs .xlsx file.

---

## Planned Features and improvements
- [ ] Add and convert the .ipynb to standalone `.py` application
- [ ] convert the code to support  odular `src/` package structure
- [ ] Add unit tests
- [ ] Docker containerization
- [ ] Code simplification, increase scalability and robustness
- [ ] Price generation improvements, more thorough product and item selection
- [ ] Integration of `streamlit` / `Flet` for mobile, web desktop usage

---

## Installation

1. **Clone the repository**
```bash
git clone https://github.com/yourusername/invoice-generator.git
cd invoice-generator
```

2. **Install dependencies**
```bash
pip install -r requirements.txt
