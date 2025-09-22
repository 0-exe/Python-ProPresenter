# ProPresenter Playlist Creator

This application helps you create a ProPresenter playlist from a Word document containing your service plan. It uses the Gemini AI to detect songs and also finds psalms in the document, then uses the ProPresenter API to create a playlist.

## Setup

### Prerequisites
- Python 3 (https://www.python.org/downloads/)
- ProPresenter 7 with the HTTP API enabled.
- A Google Gemini API Key.

### 1. Project Setup
1.  **Download or clone this project.**
2.  **Open a terminal or command prompt** and navigate to the project directory:
    ```bash
    cd /path/to/Python-ProPresenter/Python-ProPresenter
    ```
3.  **Create a virtual environment.** This keeps the project's dependencies isolated.
    ```bash
    python3 -m venv .venv
    ```
4.  **Activate the virtual environment.**
    - On macOS/Linux:
      ```bash
      source .venv/bin/activate
      ```
    - On Windows:
      ```bash
      .venv\Scripts\activate
      ```
5.  **Install the required dependencies.**
    ```bash
    pip install -r requirements.txt
    ```
    This will install `python-docx`, `requests`, `beautifulsoup4`, and `google-generativeai`.

### 2. Enable ProPresenter API
1. Open ProPresenter.
2. Go to `Preferences` > `Network`.
3. Make sure `Enable Network` is checked.
4. Under `HTTP Server`, check `Enable HTTP Server`.
5. Note the `Port` number (default is 1025). The application assumes the default port.

## How to Use

1.  **Run the application:**
    ```bash
    python gui.py
    ```
2.  The main application window will appear.

### Step-by-Step Guide

1.  **Configure Settings (First Time):**
    - The first time you run the app, you need to configure the Gemini API settings.
    - **Get a Gemini API Key:** Go to [Google AI Studio](https://aistudio.google.com/app/apikey) and create an API key.
    - **Enter the API Key:** In the "Settings" section of the application, paste your API key into the "Gemini API Key" field.
    - **Model Name:** The `gemini-pro` model is set by default and works with the free tier. You can change this to other models if you wish.
    - **Save Settings:** Click the "Save Settings" button. Your API key will be stored locally in a `config.json` file so you don't have to enter it again. This file is ignored by git.

2.  **Select Word Document:**
    - In the "Create Playlist" section, click the "Browse..." button to select the `.docx` file containing your service plan.
    - The application will automatically fill in the "Playlist Name" based on the filename, but you can change it.

3.  **Start the Process:**
    - Click the "Start" button.
    - The application will read your document, send the text to Gemini AI for song detection, and display any songs it finds in the log area.

4.  **Import Songs into ProPresenter:**
    - If the log prompts you to import songs, open ProPresenter and use CCLI SongSelect to import the listed songs into your library.
    - **This is a manual step you must perform in ProPresenter.**

5.  **Continue:**
    - Once you have imported the songs into ProPresenter, click the "Continue" button in the application.

6.  **Playlist Creation:**
    - The application will then:
        - Fetch the text for any psalms in your plan.
        - Connect to your ProPresenter library.
        - Create a new playlist.
        - Add the songs (that it finds in the library) and psalms to the new playlist.

7.  **Check the Log:**
    - The log area will show the progress and the final result. If a song from your plan is not found in your ProPresenter library, a message will be shown.

8.  **Quit:**
    - Click the "Quit" button to close the application.

## How the `.docx` file should be formatted

With the power of Gemini AI, you don't need to follow a strict format for songs anymore. Just list them naturally in your document. The AI is trained to find them.

For psalms, you should still use the format the application recognizes:

-   **Psalms:** Psalms should be mentioned with the "Psalm" keyword followed by the number.
    ```
    Lesung: Psalm 23
    ```
