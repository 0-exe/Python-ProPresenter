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
        return f"Error parsing docx file: {e}"

def get_psalm(psalm_reference, translation):
    """Fetches psalm text from a Bible API."""
    try:
        response = requests.get(f"https://bible-api.com/{psalm_reference}?translation={translation}")
        response.raise_for_status()
        data = response.json()
        return data["text"]
    except requests.exceptions.RequestException as e:
        return f"Error fetching psalm: {e}"

def get_library():
    """Fetches the ProPresenter library."""
    try:
        response = requests.get(f"{PROPPRESENTER_URL}/api/v1/library")
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return f"Error fetching ProPresenter library: {e}"

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
        return f"Error creating ProPresenter presentation: {e}"

def add_to_propresenter_playlist(playlist_id, presentation_id):
    """Adds a presentation to a playlist in ProPresenter."""
    try:
        response = requests.post(f"{PROPPRESENTER_URL}/api/v1/playlist/{playlist_id}/items", json={"id": presentation_id})
        response.raise_for_status()
        return None
    except requests.exceptions.RequestException as e:
        return f"Error adding to ProPresenter playlist: {e}"

def create_propresenter_playlist(playlist_name, items, library):
    """Creates a playlist in ProPresenter and adds items to it."""
    log = []
    try:
        response = requests.post(f"{PROPPRESENTER_URL}/api/v1/playlist", json={"name": playlist_name})
        response.raise_for_status()
        playlist_id = response.json()["id"]
        log.append(f"Successfully created playlist '{playlist_name}' with ID: {playlist_id}")

        for item in items:
            if item["type"] == "song":
                presentation = find_presentation_in_library(library, item["name"])
                if presentation:
                    error = add_to_propresenter_playlist(playlist_id, presentation["id"])
                    if error:
                        log.append(error)
                    else:
                        log.append(f"Added '{item['name']}' to the playlist.")
                else:
                    log.append(f"Could not find '{item['name']}' in the ProPresenter library.")
            elif item["type"] == "psalm":
                presentation = create_propresenter_presentation(item["name"], item["content"])
                if presentation and 'id' in presentation:
                    error = add_to_propresenter_playlist(playlist_id, presentation["id"])
                    if error:
                        log.append(error)
                    else:
                        log.append(f"Added '{item['name']}' to the playlist.")
                else:
                    log.append(f"Failed to create presentation for '{item['name']}'.")
        
        log.append(f"Playlist '{playlist_name}' created successfully.")
        return "\n".join(log)

    except requests.exceptions.RequestException as e:
        error_message = f"Error creating ProPresenter playlist: {e}"
        log.append(error_message)
        return "\n".join(log)