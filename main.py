import os
import requests
import certifi
from dotenv import load_dotenv
import urllib3
from pptx import Presentation
import datetime

# Disable insecure request warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load environment variables from .env file
load_dotenv()

# Set SSL certificate environment variables before starting HTTPS connections
os.environ["SSL_CERT_FILE"] = certifi.where()
os.environ["REQUESTS_CA_BUNDLE"] = certifi.where()

# Retrieve configuration from environment variables
tenant_id = os.getenv("TENANT_ID")
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
resource = os.getenv("RESOURCE")
deployment_name = os.getenv("DEPLOYMENT_NAME")
openai_api_base = os.getenv("OPENAI_API_BASE")
subscription_key = os.getenv("SUBSCRIPTION_KEY")

# Function to obtain an Azure AD access token
def get_access_token():
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": resource + ".default"
    }
    token_headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    response = requests.post(token_url, data=token_data, headers=token_headers)
    response.raise_for_status()
    return response.json().get("access_token")

# Function to log errors to a file
def log_error_to_file(error_message, response_text=None):
    log_folder = "error_logs"
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file = os.path.join(log_folder, f"error_{timestamp}.log")
    with open(log_file, "w", encoding="utf-8") as f:
        f.write(f"Error occurred at {timestamp}\n")
        f.write(f"Error message: {error_message}\n")
        if response_text:
            f.write(f"Response content:\n{response_text}\n")
    print(f"Error details logged to {log_file}")

# Class to handle OpenAI text generation requests
class OpenAITextGenerator:
    def __init__(self, api_base, deployment, access_token, subscription_key):
        self.api_base = api_base.rstrip("/")  # remove trailing slash if present
        self.deployment = deployment
        self.access_token = access_token
        self.subscription_key = subscription_key
        self.api_version = "2024-07-01-preview"
    
    def send_request(self, system_message, user_message):
        api_url = f"{self.api_base}/deployments/{self.deployment}/chat/completions?api-version={self.api_version}"
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Ocp-Apim-Subscription-Key": self.subscription_key,
            "Content-Type": "application/json",
            "api-key": self.access_token
        }
        data = {
            "messages": [
                {"role": "system", "content": system_message},
                {"role": "user", "content": user_message}
            ]
        }
        try:
            response = requests.post(api_url, headers=headers, json=data, verify=False)
            response.raise_for_status()  # Raise an HTTPError for bad responses (4xx and 5xx)
            result = response.json()
            return result['choices'][0]['message']['content']
        except requests.exceptions.HTTPError as http_err:
            print(f"HTTP error occurred: {http_err}")
            print(f"Response content: {response.text}")
            log_error_to_file(str(http_err), response.text)
        except requests.exceptions.RequestException as req_err:
            print(f"Request error occurred: {req_err}")
            log_error_to_file(str(req_err))
        except KeyError as key_err:
            print(f"Unexpected response format: {key_err}")
            log_error_to_file(str(key_err), response.text)
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            log_error_to_file(str(e))

# Function to list all ppt or pptx files in a given folder
def list_ppt_files(folder_path):
    ppt_files = []
    for filename in os.listdir(folder_path):
        if filename.lower().endswith((".ppt", ".pptx", ".pptm")):
            ppt_files.append(filename)
    ppt_files.sort()
    return ppt_files

# Function to extract text from a PowerPoint file
def extract_text_from_ppt(file_path):
    text = ""
    presentation = Presentation(file_path)
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    text += paragraph.text + "\n"
    return text

def main():
    ppt_folder = os.path.join("C:\\", "python_scripts", "pptChat", "ppt")
    if not os.path.exists(ppt_folder):
        print(f"Folder not found: {ppt_folder}")
        return

    # List PowerPoint files
    ppt_files = list_ppt_files(ppt_folder)
    if not ppt_files:
        print("No PowerPoint files found in the folder.")
        return

    print("Select a PowerPoint file by entering its number:")
    for idx, filename in enumerate(ppt_files, start=1):
        print(f"{idx}: {filename}")
    selected_num = int(input("Enter the number: "))
    selected_file = ppt_files[selected_num - 1]
    ppt_path = os.path.join(ppt_folder, selected_file)

    # Extract text from the selected PowerPoint file
    extracted_text = extract_text_from_ppt(ppt_path)
    if not extracted_text.strip():
        print("No text extracted from the selected PowerPoint file.")
        return

    print("Text extracted from the PowerPoint file:")
    print(extracted_text)

    # Set the extracted text as the system prompt
    system_message = extracted_text
    print("System prompt set with the extracted PowerPoint text.")

    # Instantiate the text generator
    generator = OpenAITextGenerator(openai_api_base, deployment_name, get_access_token(), subscription_key)

    # Chat with the extracted text as the system prompt
    while True:
        user_message = input("Enter your prompt (type 'exit' to quit): ")
        if user_message.strip().lower() == "exit":
            break
        response = generator.send_request(system_message, user_message)
        responseLabel = "\n\n" + "Response:" + "\n"
        print(responseLabel)
        print(response)
        print()

if __name__ == "__main__":
    main()