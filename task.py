# Step to follow 

# Required Library need to be installed
# >>> pip install pandas 
# >>> pip install openpyxl
# >>> pip install tabulate

import pandas as pd            # For adding data and read excel file
import os                      # For interact with operating system like creating excel file
import re                      # For matching patterns or validations
from tabulate import tabulate  # Creating tabular form for display users is table form

class UserManagement:
    def __init__(self, filename="users.xlsx"):
        self.filename = filename
        self.load_or_create_excel()


    def load_or_create_excel(self):
        # Create an Excel file with headers if it does not exist
        if not os.path.exists(self.filename):
            df = pd.DataFrame(columns=["Name", "Email", "Phone Number"])
            df.to_excel(self.filename, index=False)


    def add_user(self):
        # Adding a user in excel file with the validations
        name = input("Enter Name: ").strip()
        if not self.validate_name(name):
            print("Invalid name! Only letters and spaces are allowed.")
            return

        email = input("Enter Email: ").strip()
        if not self.validate_email(email):
            print("Invalid email!")
            return

        phone = input("Enter Phone Number (10 digits, starts with 6-9): ").strip()
        if not self.validate_phone(phone):
            print("Invalid phone number! Must be 10 digits and start with 6-9.")
            return

        new_user = {"Name": name, "Email": email, "Phone Number": phone}
        self.save_to_excel(new_user)
        print("User added successfully.")

    def save_to_excel(self, user_data):
        df = pd.read_excel(self.filename)
        df = df._append(user_data, ignore_index=True)
        df.to_excel(self.filename, index=False)

    def display_users(self):
        df = pd.read_excel(self.filename)
        if df.empty:
            print("No users found.")
        else:
            print("\nStored Users:")
            print(tabulate(df, headers="keys", tablefmt="grid", showindex=False))

    @staticmethod
    def validate_name(name):
        return bool(re.match(r"^[A-Za-z\s]+$", name))

    @staticmethod
    def validate_email(email):
        return bool(re.match(r"^[\w\.-]+@[\w\.-]+\.\w{2,4}$", email))

    @staticmethod
    def validate_phone(phone):
        return bool(re.match(r"^[6-9][0-9]{9}$", phone))

if __name__ == "__main__":
    user_mgmt = UserManagement()
    while True:
        print("\nUser Management System")
        print("1 - Add User\n2 - Display Users\n3 - Exit")
        choice = input("Enter Your Choice: ")

        if choice == "1":
            user_mgmt.add_user()
        elif choice == "2":
            user_mgmt.display_users()
        elif choice == "3":
            print("Exiting the application. Thank you!")
            break
        else:
            print("Invalid option. Please try again.")
