import pandas as pd

def aggregate_sentiment_scores_by_week_and_recipient(input_excel_path, output_excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Convert the 'Date' column to datetime format
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
    
    # Create a new column for Year-Week
    df['Year-Week'] = df['Date'].dt.to_period('W')
    
    # Identify sentiment score columns for aggregation
    sentiment_columns = df.columns[5:12]  # Adjust based on your DataFrame
    
    # Get a list of unique recipients
    recipients = df['Recipient'].unique()
    
    # Initialize Excel writer
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for recipient in recipients:
            # Filter data for the current recipient
            recipient_df = df[df['Recipient'] == recipient]
            
            # Ensure correct aggregation of sentiment scores by Year-Week for the current recipient
            aggregated_scores = recipient_df.groupby(['Recipient', 'Year-Week'])[sentiment_columns].sum().reset_index()
            
            # Write the aggregated scores to a new sheet named after the recipient
            sheet_name = f"{recipient[:25]}".replace(':', '').replace('\\', '').replace('/', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
            aggregated_scores.to_excel(writer, sheet_name=sheet_name, index=False)

def main():
    input_excel_path = 'messages.xlsx'
    output_excel_path = 'messages.xlsx'
    aggregate_sentiment_scores_by_week_and_recipient(input_excel_path, output_excel_path)
    print("Completed")

if __name__ == "__main__":
    main()
