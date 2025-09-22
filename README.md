# ProPresenter Playlist Creator

This application helps you create a ProPresenter playlist from a Word document containing your service plan. It parses a `.docx` file, identifies songs and psalms, and uses the ProPresenter API to create a playlist.

## Setup

### Prerequisites
- Python 3 (https://www.python.org/downloads/)
- ProPresenter 7 with the HTTP API enabled.

### 1. Enable ProPresenter API
1. Open ProPresenter.
2. Go to `Preferences` > `Network`.
3. Make sure `Enable Network` is checked.
4. Under `HTTP Server`, check `Enable HTTP Server`.
5. Note the `Port` number (default is 1025). The application assumes the default port.

### 2. Project Setup
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

## How to Use

1.  **Run the application:**
    ```bash
    python gui.py
    ```
2.  The main application window will appear.

### Step-by-Step Guide

1.  **Select Word Document:**
    - Click the "Browse..." button to select the `.docx` file containing your service plan.
    - The application will automatically fill in the "Playlist Name" based on the filename, but you can change it.

2.  **Enter Playlist Name:**
    - If you don't like the automatically filled name, enter a new name for your ProPresenter playlist.

3.  **Select Psalm Translation:**
    - Choose your preferred Bible translation for the psalms from the dropdown menu.

4.  **Start the Process:**
    - Click the "Start" button.
    - The application will parse your document and display any songs it finds in the log area.

5.  **Import Songs into ProPresenter:**
    - If the log prompts you to import songs, open ProPresenter and use CCLI SongSelect to import the listed songs into your library.
    - **This is a manual step you must perform in ProPresenter.**

6.  **Continue:**
    - Once you have imported the songs into ProPresenter, click the "Continue" button in the application.

7.  **Playlist Creation:**
    - The application will then:
        - Fetch the text for any psalms in your plan.
        - Connect to your ProPresenter library.
        - Create a new playlist.
        - Add the songs (that it finds in the library) and psalms to the new playlist.

8.  **Check the Log:**
    - The log area will show the progress and the final result. If a song from your plan is not found in your ProPresenter library, a message will be shown.

9.  **Quit:**
    - Click the "Quit" button to close the application.

## How the `.docx` file should be formatted

The application looks for specific keywords in your `.docx` file to identify songs and psalms.

-   **Songs:** Songs should be in a "Lobpreis-Block" section, with each song on a new line prefixed with a `-`.
    ```
    Lobpreis-Block
    - Amazing Grace
    - How Great Thou Art
    ```

-   **Psalms:** Psalms should be mentioned with the "Psalm" keyword followed by the number.
    ```
    Lesung: Psalm 23
    ```
