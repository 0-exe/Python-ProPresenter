import docx
import requests
import re

# ProPresenter API configuration
PROPPRESENTER_URL = "http://localhost:1025"

def parse_docx(file_path):
    """Parses the .docx file to extract songs and psalms."""
    try:
        doc = docx.Document(file_path)
        full_text = [para.text for para in doc.paragraphs]
        
        items = []
        in_praise_block = False
        
        for line in full_text:
            if "Lobpreis-Block" in line:
                in_praise_block = True
                continue
            
            if in_praise_block and line.strip().startswith("-"):
                song_title = line.strip().lstrip("-").strip()
                items.append({"type": "song", "value": song_title})
            
            if in_praise_block and not line.strip().startswith("-") and line.strip() != "":
                in_praise_block = False
            
            if "Psalm" in line:
                match = re.search(r"Psalm (\d+)", line)
                if match:
                    psalm_number = match.group(1)
                    items.append({"type": "psalm", "value": f"Psalm {psalm_number}"})
        
        return items
    except Exception as e:
        print(f"Error parsing docx file: {e}")
        return []

def get_song_lyrics(song_title):
    """Returns the song title."""
    return song_title

def get_psalm(psalm_reference, translation):
    """Fetches psalm text from a Bible API."""
    print(f"Fetching psalm: {psalm_reference} ({translation})")
    try:
        response = requests.get(f"https://bible-api.com/{psalm_reference}?translation=luther")
        response.raise_for_status()
        data = response.json()
        return data["text"]
    except requests.exceptions.RequestException as e:
        print(f"Error fetching psalm: {e}")
        return f"Could not fetch text for {psalm_reference}"

def get_library():
    """Fetches the ProPresenter library."""
    try:
        response = requests.get(f"{PROPPRESENTER_URL}/api/v1/library")
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching ProPresenter library: {e}")
        return []

def find_presentation_in_library(library, name):
    """Finds a presentation in the ProPresenter library."""
    for item in library:
        if item["name"] == name:
            return item
    return None

def create_propresenter_presentation(name, content):
    """Creates a presentation in ProPresenter."""
    try:
        presentation = {
            "name": name,
            "slides": [
                {
                    "text": content
                }
            ]
        }
        response = requests.post(f"{PROPPRESENTER_URL}/api/v1/presentation", json=presentation)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Error creating ProPresenter presentation: {e}")
        return None

def add_to_propresenter_playlist(playlist_id, presentation_id):
    """Adds a presentation to a playlist in ProPresenter."""
    try:
        response = requests.post(f"{PROPPRESENTER_URL}/api/v1/playlist/{playlist_id}/items", json={"id": presentation_id})
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error adding to ProPresenter playlist: {e}")

def create_propresenter_playlist(playlist_name, items):
    """Creates a playlist in ProPresenter and adds items to it."""
    try:
        response = requests.post(f"{PROPPRESENTER_URL}/api/v1/playlist", json={"name": playlist_name})
        response.raise_for_status()  # Raise an exception for bad status codes
        playlist_id = response.json()["id"]
        print(f"Successfully created playlist '{playlist_name}' with ID: {playlist_id}")

        library = get_library()

        for item in items:
            if item["type"] == "song":
                presentation = find_presentation_in_library(library, item["name"])
                if presentation:
                    add_to_propresenter_playlist(playlist_id, presentation["id"])
                    print(f"Added '{item['name']}' to the playlist.")
                else:
                    print(f"Could not find '{item['name']}' in the ProPresenter library.")
            elif item["type"] == "psalm":
                presentation = create_propresenter_presentation(item["name"], item["content"])
                if presentation:
                    add_to_propresenter_playlist(playlist_id, presentation["id"])
                    print(f"Added '{item['name']}' to the playlist.")

    except requests.exceptions.RequestException as e:
        print(f"Error creating ProPresenter playlist: {e}")

def main():
    """Main function to orchestrate the process."""
    # 1. Get file path from user
    file_path = "/Users/mats/Documents/Python-ProPresenter/Python-ProPresenter/example.docx"

    # 2. Parse the .docx file
    service_items = parse_docx(file_path)
    
    songs_to_import = [item["value"] for item in service_items if item["type"] == "song"]
    if songs_to_import:
        print("Please import the following songs into ProPresenter from CCLI SongSelect:")
        for song in songs_to_import:
            print(f"- {song}")
        input("Press Enter to continue after importing the songs...")

    # 3. Get lyrics and psalm text
    playlist_items = []
    for item in service_items:
        if item["type"] == "song":
            playlist_items.append({"type": "song", "name": item["value"]})
        elif item["type"] == "psalm":
            psalm_text = get_psalm(item["value"], "luther") # Using 'luther' as a fallback for 'Luther17'
            playlist_items.append({"type": "psalm", "name": item["value"], "content": psalm_text})
    
    # 4. Create ProPresenter playlist
    playlist_name = "Gottesdienst 6.Juli 2025"
    create_propresenter_playlist(playlist_name, playlist_items)

if __name__ == "__main__":
    main()
