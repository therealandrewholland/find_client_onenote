import ctypes
import functools
import logging
import json
import os
import re
import string
import sys
import time
import urllib.parse

from PIL import Image
import pyperclip
import pystray
from tendo import singleton
import tkinter as tk
from tkinter import messagebox, ttk
import keyboard
import webbrowser

class JSONFileManager:
    @staticmethod
    def load_file(filename):
        if os.path.exists(filename):
            logging.info(f"Loading {filename}.")
            with open(filename, "r") as file:
                json_file = json.load(file)
                logging.info(f"{filename} loaded.")
                return json_file
        return None

    @staticmethod
    def save_file(filename, data):
        with open(filename, "w") as file:
            logging.info(f"Saving {filename}.")
            json.dump(data, file, indent=4)
            logging.info(f"{filename} saved.")

    @staticmethod
    def init_file(filename, initial_data):
        if not os.path.exists(filename):
            logging.info(f"{filename} does not exist. Creating file with default values.")
            JSONFileManager.save_file(filename, initial_data)
            logging.info(f"{filename} created.")

class MainApplication(tk.Tk):
    IS_GUI_INIT = False
    
    APPLICATION_NAME = "Find Client OneNote"
    
    CURRENT_ONENOTE_HOTKEY = None
    CURRENT_TICKET_SEARCH_HOTKEY = None

    NOC_PLAYBOOK = "https://aunalytics.sharepoint.com/sites/ManagedServices-NOC/_layouts/OneNote.aspx?id=%2Fsites%2FManagedServices-NOC%2FShared%20Documents%2FNOC%2FNOC"
    MAIN_DIRECTORY_PATH = os.path.join(os.environ['OneDrive'], "Clients & Prospects")
    CLIENT_LIST = os.listdir(MAIN_DIRECTORY_PATH)
    UI_COLORS = {
        'Aunalytics mode': {
            'bg': '#0a4477',
            'fg': '#FFFFFF',
            'accent': '#23bddd', 
            'entry': '#ABABAB', 
        },
        'Dark mode': {
            'bg': '#1A1A1A',
            'fg': '#FFFFFF',
            'accent': '#FFA500',
            'entry': '#333333',
        },
        'Hacker mode': {
            'bg': '#000000',  
            'fg': '#149414', 
            'accent': '#2b5329',
            'entry': '#000000',
        },
        'Vaporwave mode': {
            'bg': '#1976D2',
            'fg': '#E0E0E0',
            'accent': '#E040FB',
            'entry': '#880E4F',
        }
    }

    def __init__(self):        
        super().__init__()
        logging.info("- - - - - - - - Start - - - - - - - -")
        logging.info("Application initializing...")
        
        self.settings = self.load_settings()
        self.icon = None

        logging.info("Initializing GUI...")
        self.title(self.APPLICATION_NAME)
        self.configure(bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.init_gui()
        self.protocol('WM_DELETE_WINDOW', self.withdraw_window)
        logging.info("GUI initialized...")

        self.CURRENT_ONENOTE_HOTKEY = keyboard.add_hotkey(self.settings["OneNote Shortcut"], lambda: self.after(0, self.show_window()))
        self.CURRENT_TICKET_SEARCH_HOTKEY = keyboard.add_hotkey(self.settings["Ticket Search Shortcut"], self.search_ticket)

        MainApplication.IS_GUI_INIT = True
        logging.info("Application initialized.")
        self.withdraw_window()
        
    def load_settings(self):
        settings_filename = "settings.json"
        default_settings = {
            'UI Mode': 'Aunalytics mode',
            'OneNote Shortcut': 'Alt+V',
            'Ticket Search Shortcut': 'Alt+Z',
        }
        
        JSONFileManager.init_file(settings_filename, default_settings)
        return JSONFileManager.load_file(settings_filename)

    def save_settings(self):
        settings_filename = "settings.json"

        new_hotkeys = self.change_hotkeys()
        if new_hotkeys == False: return   
        self.settings["OneNote Shortcut"] = new_hotkeys[0]
        self.settings["Ticket Search Shortcut"] = new_hotkeys[1]
        
        self.settings["UI Mode"] = self.ui_combobox.get()
        self.change_ui_color(self.UI_COLORS[self.settings["UI Mode"]])
        
        JSONFileManager.save_file(settings_filename, self.settings)   

    def load_favorites(self): 
        favorites_filename = "favorites.json"
        default_favorites = []
        
        JSONFileManager.init_file(favorites_filename, default_favorites)
        return JSONFileManager.load_file(favorites_filename)

    def save_favorites(self, new_favorites): 
        favorites_filename = "favorites.json"
        favorites = self.load_favorites()
        
        new_favorites = self.remove_text_heart(new_favorites)
  
        if new_favorites in favorites:
            favorites.remove(new_favorites)
        else:
            favorites.append(new_favorites)
        
        favorites.sort()

        JSONFileManager.save_file(favorites_filename, favorites)    

    def change_hotkeys(self):
        onenote_shortcut = self.shortcut_entry.get()
        ticket_searcher_shortcut = self.shortcut2_entry.get()

        if not self.validate_shortcut(onenote_shortcut):
            messagebox.showinfo('Invalid Shortcut Entered', '\'%s\' is not a valid shortcut. Please use the following format: Modifier+Key, e.g. Ctrl+A, Alt+A+Z.' % onenote_shortcut)
            return False
        if not self.validate_shortcut(ticket_searcher_shortcut):
            messagebox.showinfo('Invalid Shortcut Entered', '\'%s\' is not a valid shortcut. Please use the following format: Modifier+Key, e.g. Ctrl+A, Alt+A+Z.' % ticket_searcher_shortcut)
            return False

        keyboard.remove_hotkey(self.CURRENT_ONENOTE_HOTKEY)
        keyboard.remove_hotkey(self.CURRENT_TICKET_SEARCH_HOTKEY)

        self.CURRENT_ONENOTE_HOTKEY = keyboard.add_hotkey(onenote_shortcut, lambda: self.after(0, self.show_window()))
        self.CURRENT_TICKET_SEARCH_HOTKEY = keyboard.add_hotkey(ticket_searcher_shortcut, self.search_ticket)

        return [onenote_shortcut, ticket_searcher_shortcut]

    def validate_shortcut(self, shortcut):
        pattern = r'^(?:(?:Ctrl|Alt|Shift|Win)(?:\+(\w{1,3})){0,})(?:\+?(?:Ctrl|Alt|Shift|Win)(?:\+(\w{1,3})){0,})?$'
        result = re.match(pattern, shortcut)
        return bool(result)
                
    def init_gui(self):
        self.label = tk.Label(self, text='Enter client name: ', fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.label.pack()


        self.search_frame = tk.Frame(self, bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.search_frame.tag = "search_frame"
        self.entry = tk.Entry(self.search_frame, fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['entry'])
        self.entry.pack()

        self.search_button = tk.Button(self.search_frame, text='Search', command=lambda:self.change_frame('Return'), fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.search_button.pack(fill=tk.BOTH)

        self.settings_button = tk.Button(self.search_frame, text='Settings', command=lambda:self.change_frame('Escape'), fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.settings_button.pack(fill=tk.BOTH)

        self.search_frame.pack()


        self.selection_frame = tk.Frame(self, bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.selection_frame.tag = "selection_frame"

        self.listbox = tk.Listbox(self.selection_frame, fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['entry'], selectforeground=self.UI_COLORS[self.settings["UI Mode"]]['fg'], selectbackground=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH)

        self.button_frame = tk.Frame(self.selection_frame, bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.button_frame.pack(side=tk.RIGHT, fill=tk.Y)

        self.open_button = tk.Button(self.button_frame, text='Open OneNote', command=lambda: self.open_file(), fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.open_button.pack(fill=tk.BOTH)

        self.open_button2 = tk.Button(self.button_frame, text='Open Network Map', command=lambda: self.open_file(network_map=True), fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.open_button2.pack(fill=tk.BOTH)

        self.back_button = tk.Button(self.button_frame, text='Back', command=lambda: self.change_frame('Escape'), fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.back_button.pack(fill=tk.BOTH)

        self.scrollbar = tk.Scrollbar(self.selection_frame, highlightbackground=self.UI_COLORS[self.settings["UI Mode"]]['accent'], troughcolor=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.listbox.yview)


        self.settings_frame = tk.Frame(self, bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.settings_frame.tag = "settings_frame"
        
        self.ui_combobox_frame = tk.Frame(self.settings_frame, bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.ui_combobox_frame.pack()
        
        self.ui_combobox_label = tk.Label(self.ui_combobox_frame, text="Select UI Mode: ", fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.ui_combobox_label.pack(side=tk.LEFT, fill=tk.BOTH)

        self.style= ttk.Style()
        self.style.theme_use('clam')
        self.update_ui_combobox()
        
        self.shortcut_frame = tk.Frame(self.settings_frame, bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.shortcut_frame.pack()
        
        self.shortcut_label = tk.Label(self.shortcut_frame, text="Application Shortcut: ", fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.shortcut_label.pack(side=tk.LEFT, fill=tk.BOTH)
        
        self.shortcut_entry = tk.Entry(self.shortcut_frame, fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['entry'])
        self.shortcut_entry.insert(0, self.settings.get("OneNote Shortcut", "Ctrl+V"))
        self.shortcut_entry.pack(side=tk.RIGHT, fill=tk.Y)

        self.shortcut2_frame = tk.Frame(self.settings_frame, bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.shortcut2_frame.pack()
        
        self.shortcut2_label = tk.Label(self.shortcut2_frame, text="Ticket Search Shortcut: ", fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['bg'])
        self.shortcut2_label.pack(side=tk.LEFT, fill=tk.BOTH)
        
        self.shortcut2_entry = tk.Entry(self.shortcut2_frame, fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['entry'])
        self.shortcut2_entry.insert(0, self.settings.get("Ticket Search Shortcut", "Ctrl+C"))
        self.shortcut2_entry.pack(side=tk.RIGHT, fill=tk.Y)

        self.save_button = tk.Button(self.settings_frame, text='Save and Change', command=self.save_settings, fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.save_button.pack(fill=tk.BOTH)

        self.back_button2 = tk.Button(self.settings_frame, text='Back', command=lambda:self.change_frame('Escape'), fg=self.UI_COLORS[self.settings["UI Mode"]]['fg'], bg=self.UI_COLORS[self.settings["UI Mode"]]['accent'])
        self.back_button2.pack(fill=tk.BOTH)


        self.bind('<Return>', self.handle_keypress)
        self.bind('<Escape>', self.handle_keypress)

        self.listbox.bind('<Button-2>', self.handle_keypress)
        self.listbox.bind('<Button-3>', self.handle_keypress)
        self.listbox.bind('<Tab>', self.handle_keypress)
        self.listbox.bind('<Up>', self.handle_keypress)
        self.listbox.bind('<Down>', self.handle_keypress)
        self.listbox.bind("<MouseWheel>", self.handle_mousewheel)

    def change_ui_color(self, ui_option):
        self.configure(bg=ui_option['bg'])
        self.label.configure(fg=ui_option['fg'], bg=ui_option['bg'])
        
        self.search_frame.configure(bg=ui_option['bg'])
        self.entry.configure(fg=ui_option['fg'], bg=ui_option['entry'])
        self.search_button.configure(fg=ui_option['fg'], bg=ui_option['accent'])
        self.settings_button.configure(fg=ui_option['fg'], bg=ui_option['accent'])

        self.selection_frame.configure(bg=ui_option['bg'])
        self.listbox.configure(fg=ui_option['fg'], bg=ui_option['entry'], selectforeground=ui_option['fg'], selectbackground=ui_option['accent'])
        self.button_frame.configure(bg=ui_option['bg'])
        self.open_button.configure(fg=ui_option['fg'], bg=ui_option['accent'])
        self.open_button2.configure(fg=ui_option['fg'], bg=ui_option['accent'])
        self.back_button.configure(fg=ui_option['fg'], bg=ui_option['accent'])
        self.scrollbar.configure(highlightbackground=ui_option['accent'], troughcolor=ui_option['bg'])

        
        self.settings_frame.configure(bg=ui_option['bg'])
        self.ui_combobox_frame.configure(bg=ui_option['bg'])
        self.update_ui_combobox()
        self.ui_combobox_label.configure(fg=ui_option['fg'], bg=ui_option['bg'])
        self.shortcut_frame.configure(bg=ui_option['bg'])
        self.shortcut_label.configure(fg=ui_option['fg'], bg=ui_option['bg'])
        self.shortcut_entry.configure(fg=ui_option['fg'], bg=ui_option['entry'])
        self.shortcut2_frame.configure(bg=ui_option['bg'])
        self.shortcut2_label.configure(fg=ui_option['fg'], bg=ui_option['bg'])
        self.shortcut2_entry.configure(fg=ui_option['fg'], bg=ui_option['entry'])
        self.save_button.configure(fg=ui_option['fg'], bg=ui_option['accent'])
        self.back_button2.configure(fg=ui_option['fg'], bg=ui_option['accent'])

    def update_ui_combobox(self):
        self.style.configure("TCombobox", 
            fieldbackground=self.UI_COLORS[self.settings["UI Mode"]]['entry'],
            background=self.UI_COLORS[self.settings["UI Mode"]]['entry'],
            foreground=self.UI_COLORS[self.settings["UI Mode"]]['fg'])
        
        self.ui_combobox_frame.option_add("*TCombobox*Listbox*Background", self.UI_COLORS[self.settings["UI Mode"]]['entry'])
        self.ui_combobox_frame.option_add('*TCombobox*Listbox*Foreground', self.UI_COLORS[self.settings["UI Mode"]]['fg'])

        if hasattr(self, 'ui_combobox'):
            self.ui_combobox.destroy()
            
        mode_list = list(self.UI_COLORS.keys())
        self.ui_combobox = ttk.Combobox(self.ui_combobox_frame, values=mode_list)
        self.ui_combobox.set(self.settings.get("UI Mode", "Dark mode"))
        self.ui_combobox.pack(side=tk.RIGHT, fill=tk.Y)

    def place_window_at_cursor(self):
        pt = ctypes.wintypes.POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))

        self.geometry(f"+{pt.x}+{pt.y}")

    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        self.update_idletasks()
        window_width = self.winfo_reqwidth()
        window_height = self.winfo_reqheight()
        
        x_cordinate = int((screen_width/2) - (window_width/2))
        y_cordinate = int((screen_height/2) - (window_height/2))

        self.geometry(f"+{x_cordinate}+{y_cordinate}")
        
    def get_active_frame(self):
        for child in self.winfo_children():
            if isinstance(child, tk.Frame) and child.winfo_ismapped():
                return child

    def change_frame_by_tag(self, active_tag, key_pressed, event_state):
        logging.info(f"Swapping frame from {active_tag} using {key_pressed} key")
        frame_actions = {
            "search_frame": self.handle_search_frame_keys,
            "selection_frame": self.handle_selection_frame_keys,
            "settings_frame": self.handle_settings_frame_keys
        }
        action = frame_actions.get(active_tag, lambda *args: None)
        action(key_pressed, event_state)

    def handle_search_frame_keys(self, key_pressed, event_state):
        if key_pressed == 'Return' and self.focus_get() == self.entry:
            if not self.search_folders():
                return
            self.search_frame.pack_forget()
            self.transition_to_selection_frame()

        if key_pressed == 'Escape':
            self.search_frame.pack_forget()
            self.transition_to_settings_frame()

    def handle_selection_frame_keys(self, key_pressed, event_state):
        if key_pressed == 'Return' and self.focus_get() == self.listbox:
            self.selection_frame.pack_forget()

            if event_state & 4: 
                self.open_file(network_map=True)
            else:
                self.open_file(network_map=False)
        if key_pressed == 'Escape':
            self.selection_frame.pack_forget()
            self.transition_to_search_frame()

    def handle_settings_frame_keys(self, key_pressed, event_state):
        if key_pressed == 'Escape':
            self.settings_frame.pack_forget()
            settings_map = {
                self.shortcut_entry: 'OneNote Shortcut',
                self.shortcut2_entry: 'Ticket Search Shortcut'
            }

            for entry, setting in settings_map.items():
                entry.delete(0, 'end')
                entry.insert(0, self.settings[setting])

            self.transition_to_search_frame()

    def transition_to_selection_frame(self):
        self.label.config(text="Select client")
        self.selection_frame.pack()
        self.listbox.focus_set()

    def transition_to_search_frame(self):
        self.label.config(text='Enter client name: ')
        self.search_frame.pack()
        self.entry.focus_set()

    def transition_to_settings_frame(self):
        self.label.config(text="Settings")
        self.settings_frame.pack()

    def change_frame(self, key_pressed, event_state=None):
        active_frame = self.get_active_frame()
        active_tag = getattr(active_frame, 'tag', None)
        self.change_frame_by_tag(active_tag, key_pressed, event_state)

    def handle_keypress(self, event):
        key_pressed = event.keysym
        active_frame = self.get_active_frame()
        
        if key_pressed in ['Return', 'Escape']:
            self.change_frame(key_pressed, event.state)
            return "break"
        elif active_frame.tag == "selection_frame" and (key_pressed == 'Tab' or event.num in [2,3]):
            if event.y:
                self.listbox.select_clear(0, tk.END)
                self.listbox.select_set(self.listbox.nearest(event.y))

            self.save_favorites(self.listbox.get(self.listbox.curselection()[0]))
            self.search_folders()

            self.listbox.focus_set()
            return "break"
        elif active_frame == "selection_frame" and key_pressed in ['Down', 'Up']:
            self.navigate_listbox(key_pressed)
            return "break"

    def handle_mousewheel(self, event):
        self.listbox.yview_scroll(int(-1*(event.delta/120)), "units")

    def navigate_listbox(self, key_pressed):
        current_index = self.listbox.index(tk.ACTIVE)
        increment_dict = {'Up': -1, 'Down': 1}
        increment = increment_dict.get(key_pressed)
        next_index = current_index + increment
        if (-1 < next_index < self.listbox.size()):
            self.listbox.selection_clear(0, tk.END)
            self.listbox.select_set(next_index)
            self.listbox.activate(next_index)
            if not self.is_index_visible(next_index):
                self.listbox.yview_scroll(increment, "units")
        return "break"

    def is_index_visible(self, index):
        first_visible_index = int(self.listbox.index("@0,0"))
        last_visible_index = int(self.listbox.index("@0,%s" % self.listbox.winfo_height()))
        return first_visible_index <= index <= last_visible_index

    def search_folders(self, event=None):
        search_query = self.entry.get().strip()
        favorites = self.load_favorites()

        if not search_query:
            if not favorites:
                if self.get_active_frame().tag == 'selection_frame':
                    logging.info(f"No input, loading selection screen with favorites.")
                    self.change_frame_by_tag('selection_frame', 'Escape', None)
                    return False
                logging.info(f"No input or favorites")
                messagebox.showinfo('No Input', 'Please enter a client name.')
                return False
            else:
                matching_clients = []
        else:
            logging.info(f"Searching for {search_query} in client list.")
            matching_clients = self.find_matching_clients(search_query, favorites)

            if not matching_clients:
                logging.info(f"No search results or favorites found.")
                messagebox.showinfo('No Results', 'No matching clients found.')
                return False
        
        logging.info(f"Results found, populating listbox with clients and favorites")
        self.populate_listbox(matching_clients, favorites)
        
        return True

    def populate_listbox(self, clients, favorites):
        self.clear_listbox()
        sorted_clients = sorted(clients)
        display_list = self.prepend_favorites(sorted_clients, favorites)
        self.insert_into_listbox(display_list)
        self.adjust_listbox_width(display_list)

        self.listbox.update_idletasks()
        self.listbox.activate(0)
        self.listbox.select_set(0)

    def prepend_favorites(self, clients, favorites):
        if not favorites:
            return clients

        for fav in reversed(favorites):
            if fav in clients:
                clients.remove(fav)
            clients.insert(0, f"♡ {fav}")
        return clients

    def insert_into_listbox(self, clients):
        for client in clients:
            self.listbox.insert(tk.END, client)

    def adjust_listbox_width(self, clients):
        max_width = max(len(client) for client in clients)
        self.listbox.config(width=max_width)

    def clear_listbox(self):
        self.listbox.delete(0, tk.END)

    def find_matching_clients(self, query, favorites):
        if not query and not favorites:
            messagebox.showinfo('No Input', 'Please enter a client name.')
            return None

        clients = self.filter_client_list_by_query(query)
        if not clients:
            messagebox.showinfo('No Results', 'No matching folders found.')
            return None

        return clients

    def filter_client_list_by_query(self, query):
        prepared_list = self.prepare_client_list()
        normalized_query = self.simplify_string(query)
        return [client for client in prepared_list if self.match_client(client, normalized_query)]

    def prepare_client_list(self):
        return ["NOC Playbook"] + self.CLIENT_LIST if "NOC Playbook" not in self.CLIENT_LIST else self.CLIENT_LIST

    def match_client(self, client, query):
        normalized_client = self.simplify_string(client)
        for abbreviation in self.get_name_abbreviation(client):
            #if query in self.simplify_string(abbreviation) or query in normalized_client:
            if self.simplify_string(abbreviation).startswith(query) or normalized_client.startswith(query):
                return True
        return False

    def get_name_abbreviation(self, client):
        abbreviate_by_caps = ''.join(char for char in client if char.isupper())
        abbreviate_by_words = ''.join(word[0] for word in client.split())
        return [abbreviate_by_caps, abbreviate_by_words]

    def simplify_string(self, complicated_string):
        return complicated_string.translate(str.maketrans('', '', string.punctuation)).lower()

    def open_file(self, network_map=False):
        try:
            selected_folder = self.get_selected_folder()
            if selected_folder == "NOC Playbook":
                self.open_noc_playbook()
                return

            folder_path, file_list = self.locate_specific_directory(selected_folder, network_map)
            file_path = self.determine_file_path(folder_path, file_list, network_map, selected_folder)

            self.open_target(file_path)
        except Exception as e:
            messagebox.showerror('Error', f"An unexpected error occurred: {str(e)}")

    def get_selected_folder(self):
        selected_folder = self.listbox.get(self.listbox.curselection()[0])
        return self.remove_text_heart(selected_folder)

    def open_noc_playbook(self):
        webbrowser.open(self.NOC_PLAYBOOK)
        self.withdraw_window()

    def locate_specific_directory(self, selected_folder, network_map):
        folder_path = os.path.join(self.MAIN_DIRECTORY_PATH, selected_folder)
        specific_directories = self.generate_specific_directories(folder_path, network_map)

        result = self.find_directory_with_files(specific_directories)
        return result if result != (None, None) else folder_path, ''

    def generate_specific_directories(self, folder_path, network_map):
        file_type = "Network & Wireless" if network_map else "OneNote"
        return [
            os.path.join(folder_path, "Operations", "Documentation", file_type),
            os.path.join(folder_path, "Documentation", file_type)
        ]

    def find_directory_with_files(self, specific_directories):
        for specific_dir in specific_directories:
            if os.path.isdir(specific_dir):
                file_list = os.listdir(specific_dir)
                if file_list:
                    return specific_dir, file_list
        return None, None

    def determine_file_path(self, folder_path, file_list, network_map, selected_folder):
        if file_list is None:
            return self.handle_missing_files(folder_path, network_map)

        if not network_map:
            return self.find_onenote_file(file_list, folder_path, selected_folder)
        else:
            return self.get_file_path("", folder_path, selected_folder)

    def handle_missing_files(self, folder_path, network_map):
        file_type = "Network & Wireless" if network_map else "OneNote"
        messagebox.showerror(f'{file_type} files not found', f'No {file_type} files found. Click "OK" to open Client\'s OneDrive folder')
        return folder_path

    def find_onenote_file(self, file_list, folder_path, selected_folder):
        extensions = ['.url', '.one']
        for ext in extensions:
            filtered_files = [file for file in file_list if file.lower().endswith(ext)]
            if filtered_files:
                return self.get_file_path(filtered_files, folder_path, selected_folder)
        return folder_path

    def get_file_path(self, file_list, specific_directory, selected_folder):
        if len(file_list) == 1:
            return os.path.join(specific_directory, file_list[0])
        return self.select_file_from_list(file_list, specific_directory, selected_folder)

    def select_file_from_list(self, file_list, specific_directory, selected_folder):
        matching_files = [file for file in file_list if selected_folder.lower() in file.lower()]
        return os.path.join(specific_directory, matching_files[0]) if len(matching_files) == 1 else specific_directory

    def remove_text_heart(self, string):
        return string.lstrip("♡ ") if string.startswith("♡ ") else string

    def open_target(self, file_path):
        logging.info(f"Opening \"{file_path}\" based on user selection")
        os.startfile(file_path)
        self.withdraw_window()

    def init_system_tray(self):
        if self.icon or not MainApplication.IS_GUI_INIT: return
        menu = (
            pystray.MenuItem('Settings', lambda: self.show_window(settings=True)),
            pystray.MenuItem('Quit', self.exit_action),
        )
        self.icon = pystray.Icon("Find_Client_OneNote", Image.open("icon.ico"), "Find Client OneNote", menu)
        self.icon.run()

    def exit_action(self):
        if not self.icon or not MainApplication.IS_GUI_INIT: return
        logging.info("Application shutting down.")
        self.icon.stop()
        self.quit()

    def show_window(self, settings=None):
        if not self.icon or not MainApplication.IS_GUI_INIT: return
        logging.info("Hiding system tray icon and opening GUI.")
        self.icon.stop()
        self.icon = None
        if settings:
            self.change_frame_by_tag('search_frame', 'Escape', None)
            self.center_window()
        else:
            self.change_frame_by_tag('settings_frame', 'Escape', None)
            self.place_window_at_cursor()
        self.deiconify()
        self.lift()
        self.focus_force()
        if settings:
            self.settings_frame.focus()
        else:
            self.entry.focus_set()
            
    def withdraw_window(self):
        logging.info("Hiding GUI and opening system tray icon.")
        self.withdraw()
        self.after(0, self.init_system_tray) 

    def get_selected_text(self):
        original_clipboard_content = pyperclip.paste()

        keyboard.release('alt')
        time.sleep(0.5)

        keyboard.press_and_release('ctrl+c')
        time.sleep(.1)
        selected_text = pyperclip.paste()
        pyperclip.copy(original_clipboard_content)
        
        return selected_text

    def search_ticket(self):
        url_encoded_text = urllib.parse.quote(self.get_selected_text())
        ticket_url = "https://ww1.autotask.net/Mvc/ServiceDesk/TicketGridSearch.mvc/SearchByTicketNumberOrTitleOrDescription?TicketNumberOrTitleOrDescription=%s" % url_encoded_text
        webbrowser.open(ticket_url)


#Useful for debugging
logging.basicConfig(
    filename='Find Client OneNote.log',
    level=logging.ERROR, #INFO
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
def log_method_call(func):
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except Exception as e:
            logging.error(f'An error occurred in method {func.__name__}: {str(e)}')
            messagebox.showerror('Error', f"An unexpected error occurred: {str(e)}")
    return wrapper

def apply_logging_to_all_methods(cls):
    for attr in dir(cls):
        if callable(getattr(cls, attr)) and not attr.startswith("__"):
            setattr(cls, attr, log_method_call(getattr(cls, attr)))

apply_logging_to_all_methods(MainApplication)
apply_logging_to_all_methods(JSONFileManager)


if __name__ == "__main__":
    try:
        me = singleton.SingleInstance()
    except singleton.SingleInstanceException:
        messagebox.showinfo('Error', "An instance of the application is already running.")
        sys.exit(0)
        
    try:
        app = MainApplication()
        app.mainloop()
    except Exception as e:
        logging.error(f'An unhandled exception occurred: {str(e)}')
        sys.exit(f'An unexpected error occurred: {str(e)}')
