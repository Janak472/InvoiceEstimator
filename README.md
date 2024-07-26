# Invoice Estimator

An Excel-based invoice estimator with VBA functionality. This project includes buttons to add new items, parties, print, save, and more.

## VBA Code

The VBA code for each button is located in the `VBA-Code` directory. Below is a brief overview of the functionality of each button.

### AddItem

This subroutine shows the `AddItem` form.

### Addtodaysdate

This subroutine enters the current date into a specific cell.

### Addtransport

This subroutine shows the `Addtransport` form.

### ClearContentssss

This subroutine clears the content of specific ranges in the worksheet.

### Nextinvoicenumber

This subroutine increments the invoice number.

### PrintInvoiceA5

This subroutine prints the current invoice.

### Recordofinvoice

This subroutine records the invoice details.

### Refreshall

This subroutine refreshes all data in the workbook.

### Saveinvoiceexcel

This subroutine saves the current invoice as an Excel file.

### SaveInvoiceAsPdf

This subroutine saves the current invoice as a PDF file.

### Showpartydetails

This subroutine shows the `Partydetails` form.

## How to Use

1. **Open the Excel file** associated with this project.
2. **Press `Alt + F11`** to open the VBA editor.
3. **Import the `.bas` files** into the VBA editor:
    - Go to `File > Import File...`
    - Select each `.bas` file and import them into your project.
4. **Save and close the VBA editor.**
5. Use the buttons in the Excel file to execute the various subroutines.

## Important Notes

- In the `SaveInvoiceAsPdf` and `Saveinvoiceexcel` subroutines, users will need to change the path for saving the Excel and PDF files to a location on their own system.
