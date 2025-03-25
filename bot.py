import cv2  
import pytesseract  
import pandas as pd  
import os  
import glob  


pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


image_folder = "images/"
excel_file = "service_orders.xlsx"

#Functions for preprocessing images 
def preprocess_image(image_path):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  #grayscale
    gray = cv2.GaussianBlur(gray, (5, 5), 0)  
    
    #Thresholding
    processed = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2
    )
    return processed

#Function to extract text from an image
def extract_text_from_image(image_path):
    processed_image = preprocess_image(image_path)  
    extracted_text = pytesseract.image_to_string(processed_image, config='--oem 3 --psm 6')

    data_dict = {}  #Dictionary 
    lines = extracted_text.split("\n")  

    for line in lines:
        if ":" in line:  #Check for key-value pairs
            key, value = line.split(":", 1) 
            data_dict[key.strip()] = value.strip()  

    return data_dict  

#Process all images in the folder and store data in Excel
def process_images():
    all_data = [] 

    for img_path in glob.glob(os.path.join(image_folder, "*.jpg")): 
        extracted_data = extract_text_from_image(img_path)  
        if extracted_data:
            all_data.append(extracted_data) 

    if all_data:  
        df = pd.DataFrame(all_data)  

        if os.path.exists(excel_file):  
            existing_df = pd.read_excel(excel_file)  
            df = pd.concat([existing_df, df], ignore_index=True)  

        df.to_excel(excel_file, index=False)  
        print(f"✅ {len(all_data)} images processed and saved to '{excel_file}'!")
    else:
        print("No valid data extracted from images.")

#menu
def bot_menu():
    print("-------------------------------------------")
    print("\n Service Order BOT Menu:")
    print("1. Process Service Order Images")
    print("2. View All Records")
    print("3. Search for a Specific Service Order ID")
    print("4. Exit")
    print("-------------------------------------------")

#view all stored records
def view_records():
    if os.path.exists(excel_file):  
        df = pd.read_excel(excel_file)  
        print("\n📄 Service Orders Data:")
        print(df)  
    else:
        print("No records found. Please process images first.")

#Function to search Service Order ID
def search_order():
    if os.path.exists(excel_file):  
        df = pd.read_excel(excel_file)  
        search_id = input("🔍 Enter Service Order ID: ").strip()  

        if search_id in df.iloc[:, 0].astype(str).values: 
            result = df[df.iloc[:, 0].astype(str) == search_id]  
            print("\n Matching Record Found:")
            print(result)
        else:
            print("No record found for this Service Order ID.")
    else:
        print("No records found. Please process images first.")

#Run
while True:
    bot_menu()
    choice = input("Choose an option: ")

    if choice == "1":
        process_images()
    elif choice == "2":
        view_records()
    elif choice == "3":
        search_order()
    elif choice == "4":
        print("Exiting BOT. Goodbye!")
        break
    else:
        print("Invalid choice. Please try again.")
