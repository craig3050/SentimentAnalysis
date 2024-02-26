import pandas as pd

def classify_sentiment_and_write_label(input_excel_path, output_excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Function to classify sentiment based on score
    def classify_sentiment(score):
        if score < -0.995:
            return -1
        elif score > 0.995:
            return 1
        else:
            return 0
    
    # Apply the classification function to the sentiment score column
    df['Classified Sentiment'] = df['Sentiment Score'].apply(classify_sentiment)
    
    # Write the DataFrame with the new column back to an Excel file
    df.to_excel(output_excel_path, index=False)


def main():
    # Example usage:
    input_excel_path = 'messages_with_sentiment.xlsx'  # Path to your input Excel file
    output_excel_path = 'messages_with_sentiment.xlsx'  # Path for the output Excel file with the new column
    classify_sentiment_and_write_label(input_excel_path, output_excel_path)
    print("Completed")

if __name__ == "__main__":
    main()