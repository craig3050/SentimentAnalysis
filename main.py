import pandas as pd
from transformers import pipeline



def add_sentiment_scores_to_excel(input_excel_path, output_excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Initialize the sentiment analysis pipeline
    sentiment_pipeline = pipeline("sentiment-analysis")
    #sentiment_pipeline = pipeline("text-classification", model="j-hartmann/emotion-english-distilroberta-base", return_all_scores=True)

    # Initialize lists to hold sentiment labels and scores
    sentiment_labels = []
    sentiment_scores = []
    
    # Process each message
    for message in df.iloc[:, 3].tolist():  # Adjust the column index if necessary
        try:
            # Apply sentiment analysis
            result = sentiment_pipeline(message)[0]  # Assuming there's only one result per message
            label = result['label']
            score = result['score'] if label == 'POSITIVE' else -result['score']
        except ValueError:
            # Handle messages that cause errors (e.g., not strings) by setting default values
            label = 'ERROR'
            score = 0
        # Append the results to the lists
        sentiment_labels.append(label)
        sentiment_scores.append(score)
    
    # Add the lists as new columns in the DataFrame
    df['Sentiment Label'] = sentiment_labels
    df['Sentiment Score'] = sentiment_scores   

    # # Apply sentiment analysis to each message in the 4th column 
    # sentiment_scores = sentiment_pipeline(df.iloc[:, 3].tolist())  # Adjust column index if necessary
    
    # # Extract the sentiment score or label and add it to the DataFrame as a new column
    # df['Sentiment Label'] = [score['label'] for score in sentiment_scores]
    # df['Sentiment Score'] = [result['score'] if result['label'] == 'POSITIVE' else -result['score'] for result in sentiment_scores]
    
    # Write the updated DataFrame to a new Excel file
    df.to_excel(output_excel_path, index=False)

def main():
    input_excel_path = 'messages.xlsx'  # Path to your input Excel file
    output_excel_path = 'messages_with_sentiment.xlsx'  # Desired path for the output Excel file
    add_sentiment_scores_to_excel(input_excel_path, output_excel_path)

if __name__ == "__main__":
    main()

