import re

import pytesseract
from PIL import Image
from openpyxl import Workbook


# ==========================================
# ITEM / DESCRIPTION ALIASES
# ==========================================
ITEM_ALIASES = [
    "item",
    "menu item",
    "item name",
    "product",
    "product name",
    "product description",
    "description",
    "particulars",
    "details",
    "service",
    "service description"
]


# STEP 1: OCR

def extract_text_from_bill(image_path):

    image = Image.open(image_path).convert("RGB")

    text = pytesseract.image_to_string(image)

    return text


# STEP 2: Check Header

def is_item_header(line):

    line_lower = line.lower().strip()

    for alias in ITEM_ALIASES:

        if alias in line_lower:

            return True

    return False

# STEP 3: Clean Price

def clean_price(price):

    price = price.strip()

# TRy with static multiple currency
    # "$ 50.00" -> "$50.00"
    # "₹ 500.00" -> "₹500.00"
    # "€ 50.00" -> "€50.00"
    # "£ 50.00" -> "£50.00"

    price = re.sub(
        r"([₹$€£])\s+",
        r"\1",
        price
    )

    return price


# STEP 4: Extract Items

def extract_items(text):

    data = [] 

    # Serial number / item / quantity / price / total
    pattern = re.compile(
        r"^\s*"
        r"\d+\s*[\.,]?\s*"                     
        r"(.+?)\s+"                            
        r"(\d+)\s+"                           
        r"([₹$€£]?\s*[\d,]+(?:\.\d{1,2})?)\s+"  
        r"([₹$€£]?\s*[\d,]+(?:\.\d{1,2})?)"     
        r"\s*$",
        re.IGNORECASE
    )
   
    summary_words = [
        "subtotal",
        "sub total",
        "grand total",
        "total",
        "tax",
        "tax rate",
        "discount",
        "amount payable",
        "net payable",
        "balance due"
    ] 

    for line in text.splitlines():

        line = line.strip()

        if not line:

            continue
       
        if is_item_header(line):

            continue


        match = pattern.search(line)  #invoice match 

        if match:

            description = match.group(1).strip()

            quantity = match.group(2)

            rate = match.group(3)

            amount = match.group(4)

          
            description = re.sub( #description
                r"\s+",
                " ",
                description
            )        

            rate = clean_price(rate)

            amount = clean_price(amount)
        

            if description.lower() in summary_words:

                continue         

            data.append([
                description,
                int(quantity),
                rate,
                amount
            ])

    return data


def save_to_excel(data, excel_path):

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Invoice"
    
    sheet["A1"] = "Item / Description"  #Excel Header

    sheet["B1"] = "Quantity"

    sheet["C1"] = "Rate / Price"

    sheet["D1"] = "Amount"

 
    for row in data:  #add data

        sheet.append(row)

    sheet.column_dimensions["A"].width = 35

    sheet.column_dimensions["B"].width = 12

    sheet.column_dimensions["C"].width = 18

    sheet.column_dimensions["D"].width = 18

    workbook.save(excel_path)

def main():

    image_path = "bill1.jpg"

    excel_path = "invoice12.xlsx"
    
    text = extract_text_from_bill(image_path)

    print("\n= OCR TEXT =\n")

    print(text) 

    data = extract_items(text) #exact item match & print item    

    print("\n= ITEMS FOUND =\n")

    if data:

        for row in data:

            print(
                "Item / Description:",
                row[0],
                "| Quantity:",
                row[1],
                "| Rate / Price:",
                row[2],
                "| Amount:",
                row[3]
            )

    else:

        print("No items found.") #save excel in folder    

    save_to_excel(
        data,
        excel_path
    )

    print(
        "\nData successfully saved to",
        excel_path
    )

if __name__ == "__main__":

    main()