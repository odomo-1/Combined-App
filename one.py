import nltk
import os

# Set the path for NLTK data
nltk_data_path = os.path.join(os.getcwd(), "nltk_data")
nltk.data.path.append(nltk_data_path)

# Download punkt if not already available
nltk.download("punkt", download_dir=nltk_data_path)