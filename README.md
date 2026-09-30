# Restaurant Billing System

A simple **Restaurant Billing System** developed in Python using **Tkinter** for the graphical user interface and **OpenPyXL** for storing billing records in an Excel file.

## 1. Project Overview

This project provides a basic restaurant billing application with:

- Login authentication
- Dashboard
- Add new customer/bill records
- Automatic price calculation
- Excel-based data storage
- View/search records
- Update records
- Delete records with confirmation
- Clear/reset form
- Logout option

## 2. Technologies Used

- **Python 3**
- **Tkinter** – GUI
- **OpenPyXL** – Excel file handling
- **datetime** – Date and time generation
- **os** – File existence checking

## 3. Project Files

The main Python program creates and uses:

```text
Restaurant_Billing_System.py
Restaurant_Bill_Data.xlsx
```

`Restaurant_Bill_Data.xlsx` is automatically created when the program is run for the first time.

## 4. Installation Requirements

Make sure Python 3 is installed on your computer.

Install the required OpenPyXL package using:

```bash
pip install openpyxl
```

Tkinter is normally included with standard Python installations on Windows.

## 5. How to Run

1. Save the Python source code as:

```text
Restaurant_Billing_System.py
```

2. Open Command Prompt/Terminal in the project folder.

3. Run:

```bash
python Restaurant_Billing_System.py
```

4. The Login Page will appear.

## 6. Login Details

Default login credentials:

```text
Username: admin
Password: 1234
```

After successful login, the Dashboard will be displayed.

## 7. Dashboard Options

The dashboard provides four options:

### 1. Add Record

Used to add a new restaurant bill.

Enter:

- Customer Name
- Phone
- Food Item
- Quantity

The price is automatically selected according to the food item.

### 2. View Records

Displays saved billing records in a table.

Records can be searched using the customer name.

### 3. Update / Delete

A selected record can be:

- Updated
- Deleted

Deletion requires confirmation before the record is removed.

### 4. Exit

Closes the application.

## 8. Available Food Menu

The current menu and prices are:

| Food Item | Price |
|---|---:|
| Pizza | ₹150 |
| Burger | ₹100 |
| Sandwich | ₹80 |
| French Fries | ₹70 |
| Cold Drink | ₹50 |

These values can be modified in the `MENU` dictionary in the Python program.

## 9. Bill Calculation

The total bill is calculated using:

```text
Total = Quantity × Price
```

For example:

```text
Food Item = Pizza
Quantity = 2
Price = ₹150

Total = 2 × 150
      = ₹300
```

## 10. Excel Data Storage

All records are stored in:

```text
Restaurant_Bill_Data.xlsx
```

The Excel sheet is named:

```text
Restaurant Bills
```

The following fields are stored:

| Field | Description |
|---|---|
| Bill ID | Unique bill identifier |
| Customer Name | Name of customer |
| Phone | Customer phone number |
| Food Item | Selected food item |
| Quantity | Quantity ordered |
| Price | Price per item |
| Total | Total bill amount |
| Date | Bill date and time |

## 11. Main Functional Modules

### Login Module

Validates username and password before opening the dashboard.

### Dashboard Module

Provides navigation to the major functions of the system.

### Add Record Module

Accepts billing information, calculates the total, and saves the record to Excel.

### View Module

Loads records from Excel and displays them using a Tkinter Treeview.

### Search Module

Allows records to be filtered by customer name.

### Update Module

Updates customer, phone, food item, quantity, price, and total in the Excel file.

### Delete Module

Deletes the selected record after confirmation.

## 12. Important Notes

- The Excel file is created automatically if it does not already exist.
- Keep `Restaurant_Bill_Data.xlsx` in the same folder as the Python program.
- Do not manually change the Excel column structure while the application is running.
- The current login credentials are stored directly in the Python source code.
- The application is intended as a basic academic/project-level restaurant billing system.

## 13. Future Improvements

The project can be further enhanced by adding:

- Multiple food items in a single bill
- GST/tax calculation
- Discount calculation
- Printable bill/receipt
- Customer bill history
- Better password security
- Admin/user roles
- Automatic bill number generation
- Database support such as MySQL or SQLite
- Sales reports
- Daily/monthly revenue reports
- Improved responsive GUI design

## 14. Conclusion

The Restaurant Billing System demonstrates how Python GUI programming can be combined with Excel file handling to create a simple billing and record-management application.

It is suitable for learning and demonstrating:

- Python programming
- Tkinter GUI development
- Event handling
- File handling
- Excel data management
- CRUD operations
- Basic authentication

