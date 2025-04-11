# PowerPoint Chat Assistant

This project allows users to interact with the content of PowerPoint files (`.ppt`, `.pptx`, `.pptm`) using Azure OpenAI services. The script extracts text from a selected PowerPoint file and uses it as a system prompt for generating responses in a conversational format.

---

## Features

- Extract text from PowerPoint files and use it as a system prompt.
- Interact with the extracted content through a chat interface.
- Supports `.ppt`, `.pptx`, and `.pptm` file formats.
- Logs errors to a dedicated folder for debugging.

---

## Prerequisites

### 1. Install Required Libraries
Ensure the following Python libraries are installed:
- `requests`
- `certifi`
- `python-dotenv`
- `urllib3`
- `python-pptx`

You can install the required libraries using:
```bash
pip install -r requirements.txt
```

Alternatively, install them individually:
```bash
pip install requests certifi python-dotenv urllib3 python-pptx
```

### 2. Set Up Environment Variables
Create a `.env` file in the project directory and define the following variables:
```env
TENANT_ID=<your-tenant-id>
CLIENT_ID=<your-client-id>
CLIENT_SECRET=<your-client-secret>
RESOURCE=<your-resource>
DEPLOYMENT_NAME=<your-deployment-name>
OPENAI_API_BASE=<your-openai-api-base-url>
SUBSCRIPTION_KEY=<your-subscription-key>
```

### 3. Create Required Folders
Before running the script, create the following folder:
- **`ppt`**: Store PowerPoint files that you want to use for the chat.
- **`error_logs`**: Store error logs when issues occur.

You can create the folders manually or by running the following command:
```bash
mkdir ppt error_logs
```

---

## How to Use

### 1. Run the Script
Execute the script using Python:
```bash
python main.py
```

### 2. Select a PowerPoint File
1. Place the PowerPoint file you want to use in the `ppt` folder.
2. When prompted, select the file by entering its corresponding number.

### 3. Start the Chat
1. The script will extract text from the selected PowerPoint file and use it as the system prompt.
2. Enter your prompts in the terminal to interact with the extracted content.
3. Type `exit` to quit the chat.

---

## Error Handling
If an error occurs during execution, the script logs the details in the `error_logs` folder. Check the log files for debugging.

---

## Notes
- Ensure the `.env` file contains valid credentials for Azure OpenAI services.
- The script disables SSL warnings for simplicity. For production use, ensure proper SSL configurations.
```