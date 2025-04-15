# PowerPoint Chat Assistant

This project enables users to interact with the content of PowerPoint files (`.ppt`, `.pptx`, `.pptm`) using Azure OpenAI services. The script extracts text from PowerPoint files, combines it with a predefined system prompt, and uses it as a system message for generating conversational responses.

---

## Features

- **PowerPoint Text Extraction**: Extracts text, titles, and slide numbers from PowerPoint files.
- **File Metadata**: Includes the original PowerPoint file name in the extracted JSON data.
- **System Prompt Integration**: Automatically prepends a predefined prompt (`pre_paper_prompt.txt`) to the extracted content for consistent instructions to the AI.
- **Conversation Management**: Saves conversation history with timestamps in the conversation_history folder.
- **File Management**:
  - Supports `.ppt`, `.pptx`, and `.pptm` file formats.
  - Tracks PowerPoint file updates and re-extracts text only if the file is modified.
- **Error Logging**: Logs errors to a dedicated folder for debugging.
- **JSON File Handling**: Allows users to interact with previously extracted JSON files.

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
Create a .env file in the project directory and define the following variables:
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
Before running the script, ensure the following folders exist:
- **ppt**: Store PowerPoint files for text extraction.
- **ppt_json**: Store extracted JSON files.
- **text_extraction_management_files**: Store metadata for tracking PowerPoint file updates.
- **conversation_history**: Store conversation history with timestamps.
- **system_prompt**: Store the `pre_paper_prompt.txt` file containing the predefined system prompt.

You can create the folders manually or by running the following command:
```bash
mkdir ppt ppt_json text_extraction_management_files conversation_history system_prompt
```

### 4. Add Predefined System Prompt
Place a file named `pre_paper_prompt.txt` in the system_prompt folder. This file should contain the instructions that will always be prepended to the system message.

---

## How to Use

### 1. Run the Script
Execute the script using Python:
```bash
python main.py
```

### 2. Select an Option
When prompted, select one of the following options:
1. **New ppt file**: Extract text from a PowerPoint file.
2. **Existing json files**: Interact with previously extracted JSON files.
3. **Exit**: Exit the program.

### 3. New PowerPoint File Workflow
1. Place the PowerPoint file in the ppt folder.
2. Select **Option 1** from the menu.
3. Choose the desired PowerPoint file by entering its corresponding number.
4. The script will:
   - Check if the file has been modified since the last extraction.
   - Extract text, titles, and slide numbers from the file.
   - Save the extracted data as a JSON file in the ppt_json folder.
   - Update the metadata in the text_extraction_management_files folder.

### 4. Existing JSON File Workflow
1. Select **Option 2** from the menu.
2. Choose a JSON file from the ppt_json folder or select "All JSON files" to combine all JSON data.
3. The script will load the selected JSON data and use it as the system message for the chat.

### 5. Start the Chat
1. The script will combine the predefined system prompt (`pre_paper_prompt.txt`) with the extracted or selected JSON content.
2. Enter your prompts in the terminal to interact with the content.
3. Type `exit` to quit the chat.

---

## Error Handling
If an error occurs during execution, the script logs the details in the error_logs folder. Check the log files for debugging.

---

## Notes
- Ensure the .env file contains valid credentials for Azure OpenAI services.
- The script disables SSL warnings for simplicity. For production use, ensure proper SSL configurations.
- The conversation_history folder is included in the repository, but its contents are ignored by .gitignore.

---

## Example JSON Output
The extracted JSON file will have the following structure:
```json
[
    {
        "file_name": "example.pptx",
        "title": "Slide 1",
        "slide_number": 1,
        "text": "This is the content of slide 1."
    },
    {
        "file_name": "example.pptx",
        "title": "Slide 2",
        "slide_number": 2,
        "text": "This is the content of slide 2."
    }
]
```

---

## Folder Structure
Below is the expected folder structure:
```
pptChat/
├── ppt/                         # Store PowerPoint files
├── ppt_json/                    # Store extracted JSON files
├── text_extraction_management_files/  # Store metadata for file updates
├── conversation_history/        # Store conversation history with timestamps
├── system_prompt/               # Store pre_paper_prompt.txt
├── error_logs/                  # Store error logs
├── main.py                      # Main script
├── .env                         # Environment variables
└── requirements.txt             # Python dependencies
```

---

## License
This project is licensed under the MIT License. See the LICENSE file for details.
