import pandas as pd
from transformers import pipeline

def add_sentiment_scores_to_excel(input_excel_path, output_excel_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_path)
    
    # Initialize the sentiment analysis pipeline with the specified model
    sentiment_pipeline = pipeline("text-classification", model="j-hartmann/emotion-english-distilroberta-base", return_all_scores=True)

    # Initialize lists to hold individual sentiment scores
    anger_scores = []
    disgust_scores = []
    fear_scores = []
    joy_scores = []
    neutral_scores = []
    sadness_scores = []
    surprise_scores = []
    
    # Process each message
    for message in df.iloc[:, 3].tolist():  # Adjust the column index if necessary
        try:
            # Apply sentiment analysis
            results = sentiment_pipeline(message)[0]  # Assuming there's only one result per message
            
            # Create a dictionary to map each sentiment to its score for easier access
            score_dict = {result['label']: result['score'] for result in results}
            
            # Append the specific sentiment scores to their respective lists
            anger_scores.append(score_dict.get('anger', 0))
            disgust_scores.append(score_dict.get('disgust', 0))
            fear_scores.append(score_dict.get('fear', 0))
            joy_scores.append(score_dict.get('joy', 0))
            neutral_scores.append(score_dict.get('neutral', 0))
            sadness_scores.append(score_dict.get('sadness', 0))
            surprise_scores.append(score_dict.get('surprise', 0))
        except ValueError:
            # Handle messages that cause errors by setting default values
            anger_scores.append(0)
            disgust_scores.append(0)
            fear_scores.append(0)
            joy_scores.append(0)
            neutral_scores.append(0)
            sadness_scores.append(0)
            surprise_scores.append(0)
    
    # Add the sentiment scores as new columns in the DataFrame
    df['Anger Score'] = anger_scores
    df['Disgust Score'] = disgust_scores
    df['Fear Score'] = fear_scores
    df['Joy Score'] = joy_scores
    df['Neutral Score'] = neutral_scores
    df['Sadness Score'] = sadness_scores
    df['Surprise Score'] = surprise_scores

    # Write the updated DataFrame to a new Excel file
    df.to_excel(output_excel_path, index=False)

def main():
    input_excel_path = 'messages.xlsx'  # Path to your input Excel file
    output_excel_path = 'messages.xlsx'  # Desired path for the output Excel file
    add_sentiment_scores_to_excel(input_excel_path, output_excel_path)

if __name__ == "__main__":
    main()
