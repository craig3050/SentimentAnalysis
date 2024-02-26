import pandas as pd
import re

def text_messages_to_excel_from_file(input_file_path, excel_file_path):
    # Open and read the content of the messages.txt file
    with open(input_file_path, 'r', encoding='utf-8') as file:
        text_messages = file.read()
    
    # Define a regular expression pattern to match the date, time, recipient, and message
    pattern = re.compile(r'(\d{2}/\d{2}/\d{4}), (\d{2}:\d{2}) - (.*?): (.*)')
    
    # Split the text into lines and apply the regex to extract data
    lines = text_messages.strip().split('\n')
    data = [pattern.match(line).groups() for line in lines if pattern.match(line)]
    
    # Create a DataFrame
    df = pd.DataFrame(data, columns=['Date', 'Time', 'Recipient', 'Message Contents'])
    
    # Write the DataFrame to an Excel file
    df.to_excel(excel_file_path, index=False)

def main():
    # Example usage:
    input_file_path = 'messages.txt'  # Path to your text file containing messages
    excel_file_path = 'messages.xlsx'  # Desired path for the Excel file
    text_messages_to_excel_from_file(input_file_path, excel_file_path)
    print("Completed")

if __name__ == "__main__":
    main()

