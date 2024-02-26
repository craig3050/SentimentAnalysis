import pandas as pd

def aggregate_sentiment_scores_by_recipient_and_date(input_excel_path, output_excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Ensure sentiment scores are within the desired range (-1, 1)
    filtered_df = df[(df['Classified Sentiment'] >= -1) & (df['Classified Sentiment'] <= 1)]
    
    # Get a list of unique recipients
    recipients = filtered_df['Recipient'].unique()
    
    # Initialize Excel writer
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for recipient in recipients:
            # Filter data for the current recipient
            recipient_df = filtered_df[filtered_df['Recipient'] == recipient]
            
            # Aggregate sentiment scores by date for the current recipient
            aggregated_scores = recipient_df.groupby('Date')['Classified Sentiment'].sum().reset_index()
            
            # Write the aggregated scores to a new sheet named after the recipient
            # Use a valid Excel sheet name (<= 31 chars, remove invalid characters)
            sheet_name = f"{recipient[:25]}".replace(':', '').replace('\\', '').replace('/', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
            aggregated_scores.to_excel(writer, sheet_name=sheet_name, index=False)



def main():
    # Example usage:
    input_excel_path = 'messages_with_sentiment.xlsx'  # Path to your input Excel file with sentiment scores
    output_excel_path = 'messages_with_sentiment.xlsx'  # Output file
    aggregate_sentiment_scores_by_recipient_and_date(input_excel_path, output_excel_path)
    print("Completed")

if __name__ == "__main__":
    main()