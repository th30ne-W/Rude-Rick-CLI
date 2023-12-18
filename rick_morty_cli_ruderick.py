import requests
import openpyxl
from openpyxl.styles import Alignment

# ANSI escape codes for colors
class Color:
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    RESET = '\033[0m'

# Function to query for a character
def get_character_info(character_name):
    url = "https://rickandmortyapi.com/api/character/"
    params = {'name': character_name}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        return response.json()

# Function to query for an episode
def get_episode_name(episode_url):
    response = requests.get(episode_url)
    if response.status_code == 200:
        episode_data = response.json()
        return episode_data['name']
    else:
        print(f"Oh great, another error! This is beyond my control: {response.status_code}")
        return "Unknown Episode"

# Function to display character info in the console
def display_character_info(character_data):
    if not character_data or not character_data.get('results'):
        print(f"{Color.RED}This character doesn't exist. What are you trying to do, break me? HAHA Try harder or maybe try typing correctly for a change?{Color.RESET}\n")
        return False
    characters = character_data.get('results', [])
    for character in characters:
        print(f"{Color.GREEN}Name: {Color.CYAN}{character['name']}{Color.RESET}")
        print(f"{Color.GREEN}Status: {Color.CYAN}{character['status']}{Color.RESET}")
        print(f"{Color.GREEN}Species: {Color.CYAN}{character['species']}{Color.RESET}")
        print(f"{Color.GREEN}Location: {Color.CYAN}{character['location']['name']}{Color.RESET}\n")

        print(f"{Color.MAGENTA}Episodes they're messing around in:{Color.RESET}")
        print(f"{Color.YELLOW}{'Ep. No.':<8} | {'Episode Name':<40} | {'Episode URL'}{Color.RESET}")
        print(f"{Color.BLUE}-" * 100 + Color.RESET)

        for url in character['episode']:
            episode_number = url.split('/')[-1]
            episode_name = get_episode_name(url)
            print(f"{Color.WHITE}{episode_number:<8} | {episode_name:<40} | {url:<10}{Color.RESET}")
        print("\n")
    return True

# Function to export character info to an Excel sheet
def export_to_excel(character_data):
    if not character_data.get('results'):
        return

    character_name = character_data['results'][0]['name']
    filename = f"{character_name.replace(' ', '_')}_info.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Character Data"
    headers = ["Name", "Status", "Species", "Location", "Episode Number", "Episode Name"]
    ws.append(headers)

    for character in character_data.get('results', []):
        for url in character['episode']:
            episode_number = url.split('/')[-1]
            episode_name = get_episode_name(url)
            row = [character['name'], character['status'], character['species'], character['location']['name'], episode_number, episode_name]
            ws.append(row)

    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length

    wb.save(filename)
    print(f"Fine, I did your work. Data exported to {filename}\n")

# Function to handle interaction after displaying character info
def handle_post_display_interaction(character_info):
    if not character_info or not character_info.get('results'):
        return
    export_choice = input("Wanna save this to an Excel file? Like it matters. (yes/no): ").strip().lower()
    if export_choice == 'yes':
        export_to_excel(character_info)

# Interactive CLI function
def interactive_cli():
    print(f"\n{Color.MAGENTA}Welcome to the Rude Rick version of the Rick and Morty CLI App. Don't expect pleasantries.{Color.RESET}\n")
    while True:
        print(f"{Color.GREEN}Pick an option, if you can handle it:{Color.RESET}")
        print(f"{Color.CYAN}1. Look up a character - like it's gonna help.{Color.RESET}")
        print(f"{Color.CYAN}2. Exit - probably the best choice for you.{Color.RESET}")
        choice = input(f"{Color.YELLOW}Well? 1 or 2, it's not rocket science: {Color.RESET}")

        if choice == '1':
            while True:
                character_name = input(f"\n{Color.WHITE}Type the name of the character, if you can spell: {Color.RESET}")
                try:
                    character_info = get_character_info(character_name)
                    if display_character_info(character_info):
                        handle_post_display_interaction(character_info)
                        break
                except Exception as e:
                    print(f"{Color.RED}Oops, something broke: {e}{Color.RESET}\n")
                    break

        elif choice == '2':
            print(f"\n{Color.GREEN}Finally, you made a smart decision. Exiting!{Color.RESET}\n")
            break
        else:
            print(f"{Color.RED}That's not a valid option, but I didn't expect better from you.{Color.RESET}\n")

# Run the interactive CLI
if __name__ == "__main__":
    interactive_cli()
