from comtypes import client
import os

# Define the source and destination directories
SOURCE_DIRECTORY = '/path/to/your/word_documents'
DESTINATION_DIRECTORY = '/path/to/your/pdf_documents'

def convert_to_pdf(source, destination):
    # Initialize COM object
    word = client.CreateObject('Word.Application')
    for filename in os.listdir(source):
        if filename.endswith('.doc') or filename.endswith('.docx'):
            # Construct full file path
            file_path = os.path.join(source, filename)
            # Define the destination file path
            dest_file_path = os.path.join(destination, filename.replace('.docx', '.pdf').replace('.doc', '.pdf'))
            
            # Open the Word document
            doc = word.Documents.Open(file_path)
            # Save as PDF
            doc.SaveAs(dest_file_path, FileFormat=17)
            # Close the Word document
            doc.Close()
    # Quit Word Application
    word.Quit()

if __name__ == "__main__":
    # Create destination directory if it doesn't exist
    if not os.path.exists(DESTINATION_DIRECTORY):
        os.makedirs(DESTINATION_DIRECTORY)
    
    # Convert all Word documents in the source directory to PDF
    convert_to_pdf(SOURCE_DIRECTORY, DESTINATION_DIRECTORY)
    print("Conversion complete. Check the destination directory for PDF files.")
