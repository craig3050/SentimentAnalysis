import pandas as pd

def aggregate_sentiment_scores_by_date(input_excel_path, output_excel_path, sheet_name='Aggregated Sentiments'):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Ensure sentiment scores are within the desired range (-1, 1)
    filtered_df = df[(df['Classified Sentiment'] >= -1) & (df['Classified Sentiment'] <= 1)]
    
    # Aggregate sentiment scores by date
    aggregated_scores = filtered_df.groupby('Date')['Classified Sentiment'].sum().reset_index()
    
    # Write the aggregated scores to a new sheet in the Excel file
    with pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        aggregated_scores.to_excel(writer, sheet_name=sheet_name, index=False)



def main():
    # Example usage:
    input_excel_path = 'messages_with_sentiment.xlsx'  # Path to your input Excel file with sentiment scores
    output_excel_path = 'messages_with_sentiment.xlsx'  # Same file or different, if you want to keep it separate
    aggregate_sentiment_scores_by_date(input_excel_path, output_excel_path) 
    print("Completed")

if __name__ == "__main__":
    main()
