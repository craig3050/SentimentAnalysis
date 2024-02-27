import pandas as pd

def aggregate_sentiment_scores_by_week_and_recipient(input_excel_path, output_excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Convert the 'Date' column to datetime format (if not already)
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
    
    # Create a new column for Year-Week
    df['Year-Week'] = df['Date'].dt.to_period('W')
    
    # Specify the columns to sum for aggregation
    sentiment_columns = ['Anger Score', 'Disgust Score', 'Fear Score', 'Joy Score', 'Neutral Score', 'Sadness Score', 'Surprise Score']
    
    # Get a list of unique recipients
    recipients = df['Recipient'].unique()
    
    # Initialize Excel writer
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for recipient in recipients:
            # Filter data for the current recipient
            recipient_df = df[df['Recipient'] == recipient]
            
            # Aggregate sentiment scores by Year-Week for the current recipient
            # Use aggfunc to specify sum for each sentiment column
            aggregated_scores = recipient_df.groupby('Year-Week')[sentiment_columns].sum().reset_index()
            
            # Write the aggregated scores to a new sheet named after the recipient
            # Ensure sheet name is valid and does not exceed Excel's limit
            sheet_name = f"{recipient[:25]}".replace(':', '').replace('\\', '').replace('/', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
            aggregated_scores.to_excel(writer, sheet_name=sheet_name, index=False)

def main():
    input_excel_path = 'messages.xlsx'  # Path to your input Excel file with sentiment scores
    output_excel_path = 'messages.xlsx'  # Output file name
    aggregate_sentiment_scores_by_week_and_recipient(input_excel_path, output_excel_path)
    print("Aggregation Completed")

if __name__ == "__main__":
    main()
