import pandas as pd

def aggregate_sentiment_scores_by_recipient_and_date(input_excel_path, output_excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Assuming the sentiment score columns are from 5 to 11 (inclusive)
    # Adjust the column names as per your DataFrame after loading it
    sentiment_columns = df.columns[5:12]  # Adjust the range as needed based on your actual DataFrame
    
    # Get a list of unique recipients
    recipients = df['Recipient'].unique()
    
    # Initialize Excel writer
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for recipient in recipients:
            # Filter data for the current recipient
            recipient_df = df[df['Recipient'] == recipient]
            
            # Aggregate sentiment scores by date for the current recipient across the specified sentiment columns
            aggregated_scores = recipient_df.groupby('Date')[sentiment_columns].sum().reset_index()
            
            # Write the aggregated scores to a new sheet named after the recipient
            # Use a valid Excel sheet name (<= 31 chars, remove invalid characters)
            sheet_name = f"{recipient[:25]}".replace(':', '').replace('\\', '').replace('/', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
            aggregated_scores.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    # Example usage:
    input_excel_path = 'messages.xlsx'  # Path to your input Excel file with sentiment scores
    output_excel_path = 'messages.xlsx'  # Output file, changed to avoid overwriting input
    aggregate_sentiment_scores_by_recipient_and_date(input_excel_path, output_excel_path)
    print("Completed")

if __name__ == "__main__":
    main()
