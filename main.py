import pandas as pd
import re
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QTextEdit, QLabel, 
                            QScrollArea, QComboBox, QButtonGroup, QRadioButton, QGraphicsOpacityEffect, QDockWidget, QMainWindow,
                            QSpinBox, QFileDialog, QMessageBox, QInputDialog)
from PyQt5.QtGui import QTextCursor, QPixmap, QPalette, QIcon
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QMessageBox
import requests
import configparser
import os
from math import ceil
from difflib import get_close_matches, SequenceMatcher
from ast import literal_eval
import logging

# Setting up logging
logging.basicConfig(filename='app.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def write_ini_file(inventory_filename):
    """Writes the name of the inventory file to an .ini file."""
    config = configparser.ConfigParser()
    config["DEFAULT"] = {"InventoryFile": inventory_filename}
    
    with open("config.ini", "w") as configfile:
        config.write(configfile)

def read_ini_file():
    """Reads the name of the inventory file from the .ini file."""
    config = configparser.ConfigParser()
    config.read("config.ini")
    
    return config["DEFAULT"]["InventoryFile"]

# Reading data from the Excel file
file_path = "C:/Users/josep/Dropbox/Babcanec Works/Programming/pokemon/pokemon_card_data.xlsx"
df = pd.read_excel(file_path)

INVENTORY_FILE = read_ini_file()

# Extracting the printedTotal value from the 'set' column using regex
pattern = r'printedTotal=(\d+),'
df['printedTotal'] = df['set'].str.extract(pattern)[0].astype(int)

class PokemonCardApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowIcon(QIcon("pokemon.ico"))
        self.image_urls = []
        self.current_image_index = 0
        self.original_pixmap = None
        self.zoom_factor = 1.0
        self.image_cache = {}
        self.current_page = 0  # Pagination - current page
        self.page_size = 20  # Pagination - number of cards per page
        self.card_search = CardSearch(self)
        self.init_ui()
        

    def init_ui(self):
        # Create a central widget for the QMainWindow
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Main Layout
        layout = QVBoxLayout(central_widget)  # Set the layout to the central widget

        # Horizontal layout for search dropdown, input field, and search button
        search_hbox = QHBoxLayout()

        # Search Method Dropdown
        self.search_method_combo = QComboBox(self)
        self.search_method_combo.addItems(['Any', 'Set Number', 'Name', 'Pokedex'])
        self.search_method_combo.currentTextChanged.connect(self.update_input_placeholder)
        search_hbox.addWidget(self.search_method_combo)

        # Input field
        self.input_field = QLineEdit(self)
        self.input_field.returnPressed.connect(self.search_card)
        self.update_input_placeholder()
        search_hbox.addWidget(self.input_field)

        # Search button
        self.search_button = QPushButton('Search', self)
        self.search_button.setMaximumWidth(130)
        self.search_button.clicked.connect(self.initiate_search)
        search_hbox.addWidget(self.search_button)

        layout.addLayout(search_hbox)

        # Create a dock widget for the controls
        dock = QDockWidget("Controls", self)
        dock.setAllowedAreas(Qt.LeftDockWidgetArea)
        dock_widget = QWidget()
        dock.setWidget(dock_widget)
        dock_layout = QVBoxLayout()
        dock_layout.setSpacing(10)  # Adjust the spacing between items
        dock_layout.setContentsMargins(5, 5, 5, 5)  # Adjust margins (left, top, right, bottom)
        dock_widget.setLayout(dock_layout)
        dock.setMinimumWidth(150)

        # Collection buttons
        self.view_collection_button = QPushButton('View Collection', dock_widget)
        self.view_collection_button.setMaximumWidth(150)
        self.view_collection_button.clicked.connect(self.view_inventory)
        dock_layout.addWidget(self.view_collection_button)

        self.add_to_collection_button = QPushButton('Add to Collection', dock_widget)
        self.add_to_collection_button.setMaximumWidth(150)
        self.add_to_collection_button.clicked.connect(self.add_to_collection)
        dock_layout.addWidget(self.add_to_collection_button)

        # New Inventory button
        self.new_inventory_button = QPushButton('New Collection', dock_widget)
        self.new_inventory_button.setMaximumWidth(150)
        self.new_inventory_button.clicked.connect(self.create_new_inventory)
        dock_layout.addWidget(self.new_inventory_button)

        # Card type selection
        self.card_type_group = QButtonGroup(self)
        self.normal_button = QRadioButton('Normal', dock_widget)
        self.holofoil_button = QRadioButton('Holofoil', dock_widget)
        self.reverse_holofoil_button = QRadioButton('Reverse Holofoil', dock_widget)
        self.first_ed_holofoil_button = QRadioButton('1st Ed Holofoil', dock_widget)
        self.first_ed_normal_button = QRadioButton('1st Ed Normal', dock_widget)
        self.card_type_group.addButton(self.normal_button)
        self.card_type_group.addButton(self.holofoil_button)
        self.card_type_group.addButton(self.reverse_holofoil_button)
        self.card_type_group.addButton(self.first_ed_holofoil_button)
        self.card_type_group.addButton(self.first_ed_normal_button)
        dock_layout.addWidget(self.normal_button)
        dock_layout.addWidget(self.holofoil_button)
        dock_layout.addWidget(self.reverse_holofoil_button)
        dock_layout.addWidget(self.first_ed_holofoil_button)
        dock_layout.addWidget(self.first_ed_normal_button)
        self.normal_button.setChecked(True)  # Default to normal
        self.card_type_group.buttonClicked.connect(self.card_search.search_card)

        self.addDockWidget(Qt.LeftDockWidgetArea, dock)

        # Horizontal layout for Display table and Image
        hbox = QHBoxLayout()

        # Vertical layout for Display table and its pagination controls
        vbox_table = QVBoxLayout()

        # Display table
        self.display_table = QTableWidget(self)
        self.display_table.setColumnCount(8)
        self.display_table.setHorizontalHeaderLabels(['Name', 'ID', 'Series', 'Release Date', 'Market Price', 'High Price', 'Mid Price', 'Low Price'])
        self.display_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.display_table.itemDoubleClicked.connect(self.on_row_double_clicked)
        vbox_table.addWidget(self.display_table)

        # Pagination controls for the Display table
        pagination_layout = QHBoxLayout()
        self.prev_page_button = QPushButton('Previous Page', self)
        self.prev_page_button.clicked.connect(self.prev_page)
        self.next_page_button = QPushButton('Next Page', self)
        self.next_page_button.clicked.connect(self.next_page)
        pagination_layout.addWidget(self.prev_page_button)
        pagination_layout.addWidget(self.next_page_button)
        vbox_table.addLayout(pagination_layout)

        hbox.addLayout(vbox_table,5)
        dock_layout.addStretch(1)  # This will push all the buttons and controls to the top


        # Vertical layout for Image Viewer (Image and Navigation)
        vbox = QVBoxLayout()

        # Image label inside a scroll area
        self.image_label = QLabel(self)
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidget(self.image_label)
        self.scroll_area.setWidgetResizable(True)
        vbox.addWidget(self.scroll_area)

        # Zoom controls
        zoom_layout = QHBoxLayout()
        self.zoom_in_button = QPushButton('+', self)
        self.zoom_in_button.clicked.connect(self.zoom_in)
        self.zoom_in_button.setFixedWidth(40)
        self.zoom_out_button = QPushButton('-', self)
        self.zoom_out_button.clicked.connect(self.zoom_out)
        self.zoom_out_button.setFixedWidth(40)
        zoom_layout.addWidget(self.zoom_in_button)
        zoom_layout.addWidget(self.zoom_out_button)
        vbox.addLayout(zoom_layout)

        # Image navigation
        nav_layout = QHBoxLayout()
        self.prev_button = QPushButton('Previous Image', self)
        self.prev_button.clicked.connect(self.prev_image)
        self.next_button = QPushButton('Next Image', self)
        self.next_button.clicked.connect(self.next_image)
        nav_layout.addWidget(self.prev_button)
        nav_layout.addWidget(self.next_button)
        vbox.addLayout(nav_layout)

        hbox.addLayout(vbox,2)

        layout.addLayout(hbox)

        # Create a message label for displaying temporary fading messages
        self.message_label = QLabel(self)
        self.message_label.setAlignment(Qt.AlignCenter)
        self.message_label.setStyleSheet("background-color: rgba(255, 255, 255, 150); border: 1px solid black; padding: 10px;")
        self.message_label.hide()

        self.opacity_effect = QGraphicsOpacityEffect(self.message_label)
        self.message_label.setGraphicsEffect(self.opacity_effect)

        self.fade_timer = QTimer(self)
        self.fade_timer.timeout.connect(self.fade_out)

        self.setWindowTitle('Pokemon Card Search')
        self.resize(1400, 800)

    def update_input_placeholder(self):
        search_method = self.search_method_combo.currentText()
        if search_method == 'Set Number':
            self.input_field.setPlaceholderText("Enter card information in the format #/# or just #")
        elif search_method == 'Name':
            self.input_field.setPlaceholderText("Enter name")
        elif search_method == 'Any':
            self.input_field.setPlaceholderText("Enter name or set number")
        else:  # Pokedex
            self.input_field.setPlaceholderText("Enter Pokedex #")

    def search_card(self):
        # Delegate the search functionality to the CardSearch instance
        self.card_search.search_card()

    def on_row_double_clicked(self, item):
        # Slot to handle double-clicking a row in the table
        self.current_image_index = item.row()
        self.update_image()

    def update_image(self):
        if self.image_urls:
            image_url = self.image_urls[self.current_image_index]
            
            # Check if the image is in the cache
            if image_url in self.image_cache:
                self.original_pixmap = self.image_cache[image_url]
            else:
                response = requests.get(image_url)
                self.original_pixmap = QPixmap()
                self.original_pixmap.loadFromData(response.content)
                
                # Cache the downloaded image
                self.image_cache[image_url] = self.original_pixmap

            # Set the zoom factor such that the image fits within the scroll area by default
            # Subtract a small factor to account for potential padding or margins
            padding_factor = 0.95
            x_ratio = (self.scroll_area.width() / self.original_pixmap.width()) * padding_factor
            y_ratio = (self.scroll_area.height() / self.original_pixmap.height()) * padding_factor
            self.zoom_factor = min(x_ratio, y_ratio)

            self.apply_zoom()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Space:
            self.add_to_collection_button.click()
        super().keyPressEvent(event)

    def apply_zoom(self):
        if self.original_pixmap:
            new_size = self.original_pixmap.size() * self.zoom_factor
            scaled_pixmap = self.original_pixmap.scaled(new_size, Qt.KeepAspectRatio)
            self.image_label.setPixmap(scaled_pixmap)

    def initiate_search(self):
        self.current_page = 0
        self.search_card()

    def next_page(self):
        self.current_page += 1
        self.search_card()

    def prev_page(self):
        self.current_page = max(0, self.current_page - 1)
        self.search_card()

    def zoom_in(self):
        # Scales the current image to 10% larger based on the current zoom factor
        self.zoom_factor += 0.1
        self.apply_zoom()

    def zoom_out(self):
        # Scales the current image to 10% smaller based on the current zoom factor
        self.zoom_factor = max(0.1, self.zoom_factor - 0.1)  # Ensure we don't go below 10% of original size
        self.apply_zoom()

    def show_fading_message(self, message, duration=2000):
        """
        Display a message that fades out after a given duration.
        """
        # Reset the label's style
        self.message_label.setStyleSheet("background-color: rgba(255, 255, 255, 150); border: 1px solid black; padding: 10px;")
        self.message_label.setText(message)
        
        # Position the label at the center of the main window
        self.message_label.adjustSize()  # Adjust size to fit the text
        label_x = (self.width() - self.message_label.width()) // 2
        label_y = (self.height() - self.message_label.height()) // 2
        self.message_label.move(label_x, label_y)

        self.opacity_effect.setOpacity(1)
        self.message_label.show()

        # Start the fade-out timer
        self.fade_timer.start(50)

        # QTimer to hide the message label after the duration
        QTimer.singleShot(duration, lambda: (self.fade_timer.stop(), self.message_label.hide()))

    def fade_out(self):
        """
        Gradually decrease the opacity of the message label.
        """
        current_opacity = self.opacity_effect.opacity()
        # Decrease opacity
        if current_opacity > 0:
            self.opacity_effect.setOpacity(current_opacity - 0.05)
        else:
            self.fade_timer.stop()

    def add_to_collection(self):
        if self.image_urls:  # Ensure a card is currently displayed
            selected_row = self.display_table.currentRow()
            
            # Check if a row is selected and the cell contains a valid item
            if selected_row != -1 and self.display_table.item(selected_row, 1):
                card_id = self.display_table.item(selected_row, 1).text()  # Get the ID from the second column of the table
                
                if card_id:
                    # Extract the card details from the search results table
                    card_details = {
                        'Name': self.display_table.item(selected_row, 0).text(),
                        'ID': card_id,
                        'Series': self.display_table.item(selected_row, 2).text(),
                        'Release Date': self.display_table.item(selected_row, 3).text(),
                        'Market Price': self.display_table.item(selected_row, 4).text(),
                        'High Price': self.display_table.item(selected_row, 5).text(),
                        'Mid Price': self.display_table.item(selected_row, 6).text(),
                        'Low Price': self.display_table.item(selected_row, 7).text(),
                        'Card Type': self.card_type_group.checkedButton().text()
                    }
                    
                    # Check if card + card type combo is valid
                    if card_details['Market Price'] == '-' and \
                       card_details['High Price'] == '-' and \
                       card_details['Mid Price'] == '-' and \
                       card_details['Low Price'] == '-':
                        self.show_fading_message(f"Card {card_details['Card Type']} does not exist.", 3000)
                        self.message_label.setStyleSheet("background-color: red; border: 1px solid black; padding: 10px;")
                        return

                    # Convert to DataFrame for easier handling
                    card_df = pd.DataFrame([card_details])
                    
                    if os.path.exists(INVENTORY_FILE):
                        inventory = pd.read_excel(INVENTORY_FILE)
                        
                        # Handle case if 'ID' column and 'Card Type' doesn't exist in the inventory file
                        if 'ID' not in inventory.columns or 'Card Type' not in inventory.columns:
                            inventory = pd.DataFrame(columns=card_details.keys())
                        
                        # Check if the card already exists in the inventory with the specified card type
                        existing_card = inventory[(inventory['ID'] == card_id) & (inventory['Card Type'] == card_details['Card Type'])]
                        if not existing_card.empty:
                            # If card exists, increase the count
                            index = existing_card.index[0]
                            inventory.at[index, 'Count'] += 1
                            self.show_fading_message('Card count increased in collection.')
                        else:
                            card_df['Count'] = 1
                            inventory = pd.concat([inventory, card_df])
                            self.show_fading_message('Card added to collection.')
                        
                        inventory.to_excel(INVENTORY_FILE, index=False)
                    else:
                        card_df['Count'] = 1
                        card_df.to_excel(INVENTORY_FILE, index=False)
                        self.show_fading_message('Card added to collection.')
                else:
                    self.show_fading_message('Card ID extraction failed. Try again.')
            else:
                self.show_fading_message('Please select a card from the table first.')


    def highlight_current_card(self):
        cursor = self.display_area.textCursor()
        cursor.movePosition(QTextCursor.Start)
        cursor.movePosition(QTextCursor.Down, QTextCursor.MoveAnchor, self.current_image_index * 5)  # 5 lines per card (4 lines of text + 1 empty line)
        cursor.movePosition(QTextCursor.Down, QTextCursor.KeepAnchor, 4)  # Highlight all 4 lines
        self.display_area.setTextCursor(cursor)

    def view_inventory(self):
        inventory_path = read_ini_file()
        if not os.path.exists(inventory_path):
            choice, ok = QInputDialog.getItem(self, "Inventory Selection", "Do you want to:", ["Select an existing inventory file", "Create a new inventory file"], 0, False)
            if not ok:
                return
            if choice == "Select an existing inventory file":
                options = QFileDialog.Options()
                filePath, _ = QFileDialog.getOpenFileName(self, "Select Inventory File", "", "Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*)", options=options)
                if filePath:
                    write_ini_file(filePath)
                else:
                    QMessageBox.warning(self, "No Inventory File", "Please select a valid inventory file.")
                    return
            elif choice == "Create a new inventory file":
                file_name, ok = QInputDialog.getText(self, "New Inventory File", "Enter the name for the new inventory file (without extension):")
                if ok and file_name:
                    # Assuming you want to create an Excel file, you can modify to support CSV as well
                    file_name_with_extension = file_name + ".xlsx"
                    self.create_new_inventory(file_name_with_extension)
                    write_ini_file(file_name_with_extension)
                else:
                    QMessageBox.warning(self, "Invalid File Name", "Please provide a valid name for the new inventory file.")
                    return
            else:
                return

        self.collection_window = InventoryWindow(self)
        self.collection_window.load_inventory()
        self.collection_window.show()

    def create_new_inventory(self):
        # Ask the user for the file name and location
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getSaveFileName(self, "Save New Inventory File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
        
        # Check if a valid file name was provided
        if not file_name:
            QMessageBox.warning(self, "Invalid File Name", "Please provide a valid name for the new inventory file.")
            return

        # Ensure the file name ends with .xlsx
        if not file_name.endswith('.xlsx'):
            file_name += '.xlsx'

        # Create the new inventory file
        template_data = [
            ["Name", "ID", "Series", "Release Date", "Market Price", "High Price", "Mid Price", "Low Price", "Card Type", "Count"]
        ]
        df = pd.DataFrame(template_data)
        
        try:
            df.to_excel(file_name, index=False)
            QMessageBox.information(self, "Success", f"New inventory created at {file_name}.")
            
            # After successfully creating the inventory, update the .ini file with the path to this inventory.
            write_ini_file(file_name)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create new inventory. Error: {str(e)}")

    def prev_image(self):
        if self.image_urls:
            # If it's the first image, go to the last one
            if self.current_image_index == 0:
                self.current_image_index = len(self.image_urls) - 1
            else:
                self.current_image_index -= 1

            self.update_image()
            self.display_table.selectRow(self.current_image_index)

    def next_image(self):
        if self.image_urls:
            # If it's the last image, go to the first one
            if self.current_image_index == len(self.image_urls) - 1:
                self.current_image_index = 0
            else:
                self.current_image_index += 1

            self.update_image()
            self.display_table.selectRow(self.current_image_index)

class CardSearch:
    def __init__(self, parent):
        # Reference to the main app (PokemonCardApp)
        self.app = parent

    @staticmethod
    def custom_parser(tcgplayer_str):
        # Default data structure
        default_data = {
            'url': None,
            'updatedAt': None,
            'prices': {
                'normal': {'low': '-', 'mid': '-', 'high': '-', 'market': '-', 'directLow': '-'},
                'holofoil': {'low': '-', 'mid': '-', 'high': '-', 'market': '-', 'directLow': '-'},
                'reverseHolofoil': {'low': '-', 'mid': '-', 'high': '-', 'market': '-', 'directLow': '-'},
                'firstEditionHolofoil': {'low': '-', 'mid': '-', 'high': '-', 'market': '-', 'directLow': '-'},
                'firstEditionNormal': {'low': '-', 'mid': '-', 'high': '-', 'market': '-', 'directLow': '-'}
            }
        }

        # Check if tcgplayer_str is not a string or is blank
        if not isinstance(tcgplayer_str, str) or tcgplayer_str.strip() == "":
            no_data = default_data
            for card_type, price_data in no_data['prices'].items():
                for key in price_data:
                    price_data[key] = 'no data'
            return no_data

        # Extract URL
        url_pattern = r"url='(.*?)'"
        match_url = re.search(url_pattern, tcgplayer_str)
        url = match_url.group(1) if match_url else None

        # Extract updatedAt
        updated_pattern = r"updatedAt='(.*?)'"
        match_updated_at = re.search(updated_pattern, tcgplayer_str)
        updated_at = match_updated_at.group(1) if match_updated_at else None

        # Helper function to extract price details
        def extract_price(price_str):
            patterns = {
                'low': r"low=(\d+\.\d+)?",
                'mid': r"mid=(\d+\.\d+)?",
                'high': r"high=(\d+\.\d+)?",
                'market': r"market=(\d+\.\d+)?",
                'directLow': r"directLow=(\d+\.\d+|None)?"
            }

            extracted_prices = {}
            for key, pattern in patterns.items():
                match = re.search(pattern, price_str)
                extracted_prices[key] = match.group(1) if match and match.group(1) != "None" else "-"

            return extracted_prices

        prices = {}
        for card_type in ['normal', 'holofoil', 'reverseHolofoil', 'firstEditionHolofoil', 'firstEditionNormal']:
            pattern = f"{card_type}\s*=\s*TCGPrice\((.*?)\)"
            match = re.search(pattern, tcgplayer_str)
            if match:
                price_str = match.group(1)
                prices[card_type] = extract_price(price_str)
            else:
                prices[card_type] = default_data['prices'][card_type]

        # Check if all prices across all categories are '-'
        no_data_for_all_categories = all(
            all(price == '-' for price in price_data.values())
            for price_data in prices.values()
        )

        # If no data for all categories, replace '-' with 'no data'
        if no_data_for_all_categories:
            for price_data in prices.values():
                for key in price_data:
                    price_data[key] = 'no data'

        return {
            'url': url,
            'updatedAt': updated_at,
            'prices': prices
        }


    def similar_name(self, input_name, names_list, n=10):
        # Start with an empty list for matches
        
        # Check for an exact match
        exact_matches = [name for name in names_list if input_name.lower() in name.lower()]
        
        # Get close matches using difflib if no exact matches are found
        if not exact_matches:
            close_matches = get_close_matches(input_name.lower(), [name.lower() for name in names_list], n=n)
            close_matches = [name for name in names_list if name.lower() in close_matches]
        else:
            close_matches = []

        # Combine exact matches and close matches
        matches = exact_matches + close_matches
        
        return matches

    def extract_market_price_for_holofoil(self, tcgplayer_str):
        pattern = r"holofoil=TCGPrice\(.*?market=(\d+\.\d+)"
        match = re.search(pattern, tcgplayer_str)
        return float(match.group(1)) if match else None

    def sort_cards(self, card, input_str):
        # Create a SequenceMatcher object
        seq_matcher = SequenceMatcher(None, card['name'], input_str)

        # Get the similarity ratio
        similarity_ratio = seq_matcher.ratio()
        
        # Exact match gets highest score
        if card['name'] == input_str:
            name_score = 1000
        else:
            # Use similarity ratio as the score, but penalize names longer than the input
            name_score = similarity_ratio - 0.01 * (len(card['name']) - len(input_str))
        
        # Extracting set name for tertiary sorting
        set_name_pattern = r"name=(['\"])(.*?)\1(?=[, ])"
        match = re.search(set_name_pattern, card['set'])
        card_set_name = match.group(2) if match else "Unknown Set"
        
        # Extracting release date for secondary sorting and convert it to a sortable format
        release_date_pattern = r"releaseDate='(.*?)'"
        match = re.search(release_date_pattern, card['set'])
        if match:
            year, month, day = match.group(1).split("/")
            sortable_date = year + month + day
        else:
            # Default to an old date if not found
            sortable_date = "20000101"

        # Return a tuple (name_score, -int(sortable_date), card_set_name) for sorting
        return (name_score, int(sortable_date), card_set_name)

    def search_card(self):
        # Resetting image URLs and current image index
        self.app.image_urls = []
        self.app.current_image_index = 0

        # Getting the search input
        logging.info('Getting the search input.')
        input_str = self.app.input_field.text()
        logging.debug(f'Search input: {input_str}')

        # Searching by name
        logging.info('Searching by name.')
        names_list = df['name'].unique().tolist()
        similar_names = self.similar_name(input_str, names_list, n=10)
        name_cards_df = df[df['name'].isin(similar_names)]

        # Searching by set number
        logging.info('Searching by set number.')
        card_number_str = re.search(r"(\d+)", input_str)
        if card_number_str:
            card_number_str = card_number_str.group(1)
            set_cards_df = df[df['number'].astype(str) == card_number_str]
        else:
            set_cards_df = pd.DataFrame()

        # Filtering by set number format, e.g., '1/132'
        if re.match(r"^\d+\s*/\s*\d+$", input_str):
            set_number, total_set_number = map(int, re.split(r'\s*/\s*', input_str))
            set_cards_df = set_cards_df[set_cards_df['printedTotal'] == total_set_number]

        # Combining results from name search and set search
        combined_cards_df = pd.concat([name_cards_df, set_cards_df])
        logging.debug(f'Number of combined cards: {len(combined_cards_df)}')
        combined_cards_df.drop_duplicates(inplace=True)
        cards = sorted(combined_cards_df.to_dict(orient='records'), key=lambda card: self.sort_cards(card, input_str), reverse=True)
        logging.debug('Sorting combined cards.')

        # Implementing pagination
        start_index = self.app.page_size * self.app.current_page
        end_index = start_index + self.app.page_size
        cards = cards[start_index:end_index]

        # Convert button text to attribute name
        btn_text_mapping = {
            'normal': 'normal',
            'holofoil': 'holofoil',
            'reverse holofoil': 'reverseHolofoil',
            '1st ed holofoil': 'firstEditionHolofoil',
            '1st ed normal': 'firstEditionNormal'
        }
        selected_card_type = btn_text_mapping.get(self.app.card_type_group.checkedButton().text().lower(), 'normal')

        # Displaying the results
        if cards:
            self.app.display_table.setRowCount(len(cards))
            for index, card in enumerate(cards):
                card_name = card['name']
                card_id = card['id']
                set_name_pattern = r"name=(['\"])(.*?)\1(?=[, ])"
                match = re.search(set_name_pattern, card['set'])
                card_set_name = match.group(2) if match else "Unknown Set"

                release_date_pattern = r"releaseDate='(.*?)'"
                match = re.search(release_date_pattern, card['set'])
                release_date = match.group(1) if match else "Unknown Date"

                image_url_pattern = r"large='(.*?)'"
                match = re.search(image_url_pattern, card['images'])
                if match:
                    image_url = match.group(1)
                    self.app.image_urls.append(image_url)
                else:
                    # If the specific regex fails, try a more general approach
                    general_image_url_pattern = r"large=.*?'(https://.*?\.png)'"
                    gen_match = re.search(general_image_url_pattern, card['images'])
                    if gen_match:
                        image_url = gen_match.group(1)
                        self.app.image_urls.append(image_url)

                tcgplayer_data = CardSearch.custom_parser(card['tcgplayer'])
                prices = tcgplayer_data['prices']
                pricing = prices.get(selected_card_type, {})
                market_price = pricing.get('market', "-")
                high_price = pricing.get('high', "-")
                mid_price = pricing.get('mid', "-")
                low_price = pricing.get('low', "-")
                    
                # Setting the items for the table
                self.app.display_table.setItem(index, 0, QTableWidgetItem(card_name))
                self.app.display_table.setItem(index, 1, QTableWidgetItem(card_id))
                self.app.display_table.setItem(index, 2, QTableWidgetItem(card_set_name))
                self.app.display_table.setItem(index, 3, QTableWidgetItem(release_date))
                self.app.display_table.setItem(index, 4, QTableWidgetItem(str(market_price)))
                self.app.display_table.setItem(index, 5, QTableWidgetItem(str(high_price)))
                self.app.display_table.setItem(index, 6, QTableWidgetItem(str(mid_price)))
                self.app.display_table.setItem(index, 7, QTableWidgetItem(str(low_price)))

            self.app.update_image()
            logging.info('Updating the image.')

        else:
            # If no results are found
            self.app.display_table.setRowCount(0)
            logging.warning('No cards found.')
            QMessageBox.information(self.app, 'Information', 'Card not found.')
            self.app.image_label.clear()

class InventoryWindow(QMainWindow):
    def __init__(self, parent_app=None):
        super(InventoryWindow, self).__init__()
        self.parent_app = parent_app
        
        # Set window attributes
        self.setWindowIcon(QIcon("pokemon.ico"))
        self.setWindowTitle('Card Collection')

        # Define the action log
        self.action_log = []

        # Create a central widget for the main content and its layout
        central_widget = QWidget(self)
        layout = QVBoxLayout(central_widget)
        
        # Table widget to display the cards
        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Name', 'ID', 'Set Name'])
        layout.addWidget(self.table)

        # Set the central widget
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Create the undo dock and its contents
        self.undo_dock = QDockWidget("Undo Actions", self)
        self.undo_dock.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        
        undo_button = QPushButton("Undo", self.undo_dock)
        undo_button.clicked.connect(self.undo_last_action)
        self.undo_dock.setWidget(undo_button)
        self.undo_dock.setMinimumWidth(150)

        
        # Add the undo dock to the main window on the right
        self.addDockWidget(Qt.RightDockWidgetArea, self.undo_dock)


        # Load the inventory from the Excel file
        if os.path.exists(INVENTORY_FILE):
            self.inventory = pd.read_excel(INVENTORY_FILE)
        else:
            self.inventory = pd.DataFrame()

        # Set the default size for the window
        self.resize(1200, 500)

    def load_inventory(self):
        # Clear the table first
        self.table.setRowCount(0)
        
        # Set the columns to match the DataFrame's columns
        columns = ['Name', 'ID', 'Series', 'Release Date', 'Market Price', 'High Price', 'Mid Price', 'Low Price', 'Card Type', 'Count']

        # Check if the required columns exist in the DataFrame
        missing_columns = [col for col in columns if col not in self.inventory.columns]
        if missing_columns:
            # Create default columns if they're missing
            for col in missing_columns:
                self.inventory[col] = ""

        # Set the table column count and headers
        self.table.setColumnCount(len(columns))
        self.table.setHorizontalHeaderLabels(columns)

        # Filter the inventory to only consider the required columns
        self.inventory = self.inventory[columns]

        # Load data from the inventory DataFrame
        for index, row in self.inventory.iterrows():
            self.table.insertRow(index)
            for col, value in enumerate(row):
                self.table.setItem(index, col, QTableWidgetItem(str(value)))

        # Resize columns to fit content
        self.table.resizeColumnsToContents()

        # Adding a delete button for each row
        delete_btn_column = self.table.columnCount()
        self.table.setColumnCount(delete_btn_column + 1)
        self.table.setHorizontalHeaderItem(delete_btn_column, QTableWidgetItem("Delete"))
        
        for index in range(self.table.rowCount()):
            delete_btn = QPushButton("Delete")
            delete_btn.clicked.connect(lambda _, idx=index: self.delete_row(idx))
            self.table.setCellWidget(index, delete_btn_column, delete_btn)

        # Adding uptick and downtick buttons for each row
        uptick_btn_column = self.table.columnCount()
        self.table.setColumnCount(uptick_btn_column + 1)
        self.table.setHorizontalHeaderItem(uptick_btn_column, QTableWidgetItem("Add"))
            
        downtick_btn_column = self.table.columnCount()
        self.table.setColumnCount(downtick_btn_column + 1)
        self.table.setHorizontalHeaderItem(downtick_btn_column, QTableWidgetItem("Subtract"))
            
        for index in range(self.table.rowCount()):
            # Add button
            add_btn = QPushButton("+1")
            add_btn.clicked.connect(lambda _, idx=index: self.add_to_count(idx))
            self.table.setCellWidget(index, uptick_btn_column, add_btn)
                
            # Subtract button
            subtract_btn = QPushButton("-1")
            subtract_btn.clicked.connect(lambda _, idx=index: self.subtract_from_count(idx))
            self.table.setCellWidget(index, downtick_btn_column, subtract_btn)

    def delete_row(self, index):
        card_name = self.table.item(index, 0).text() if self.table.item(index, 0) else "Unknown Card"
        card_type = self.table.item(index, 8).text() if self.table.item(index, 8) else "Unknown Type"

        # Capture the data for the action log before deleting the row
        card_data = {
            "action": "delete",
            "data": self.inventory.iloc[index].to_dict()
        }
        
        # Now, it's safe to remove the row
        self.table.removeRow(index)
        
        self.inventory.drop(index, inplace=True)
        self.inventory.reset_index(drop=True, inplace=True)  # Important to reset index after deletion
        self.inventory.to_excel(INVENTORY_FILE, index=False)
        
        if self.parent_app:
            self.parent_app.show_fading_message(f"{card_name} ({card_type}) removed from collection.")

        self.action_log.append(card_data)

    def undo_last_action(self):
        if not self.action_log:
            QMessageBox.warning(self, "Undo", "No actions to undo!")
            return

        last_action = self.action_log.pop()
        if last_action["action"] == "delete":
            # If the last action was a delete, add the card back to the inventory
            card_data = last_action["data"]
            # Using concat instead of append
            self.inventory = pd.concat([self.inventory, pd.DataFrame([card_data])], ignore_index=True)
            self.inventory.to_excel(INVENTORY_FILE, index=False)
            self.load_inventory()
            if self.parent_app:
                self.parent_app.show_fading_message(f"Undo: {card_data['Name']} ({card_data['Card Type']}) added back to collection.")

        elif last_action["action"] in ["add", "subtract"]:
            index = last_action["index"]
            previous_count = last_action["previous_count"]
            self.table.item(index, 9).setText(str(previous_count))
            self.update_inventory_file()

    def add_to_count(self, index):
        # Increment the count by 1
        current_count = int(self.table.item(index, 9).text())
        self.table.item(index, 9).setText(str(current_count + 1))
        self.update_inventory_file()

        # Log the addition action
        card_data = {
            "action": "add",
            "index": index,
            "previous_count": int(self.table.item(index, 9).text()) - 1  # we subtract 1 because we've already increased the count
        }
        self.action_log.append(card_data)

    def subtract_from_count(self, index):
        # Decrement the count by 1. If count becomes 0, delete the row
        item = self.table.item(index, 9)
        current_count = int(item.text()) if item else 0
        if current_count > 1:
            self.table.item(index, 9).setText(str(current_count - 1))
            self.update_inventory_file()
            
            # Log the subtraction action
            card_data = {
                "action": "subtract",
                "index": index,
                "previous_count": current_count  # we don't need to adjust the count here, since we've already captured it
            }
            self.action_log.append(card_data)
        else:
            self.delete_row(index)

    def update_inventory_file(self):
        # Update the inventory DataFrame from the table widget
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount() - 3):  # Excluding last 3 columns (buttons)
                self.inventory.iat[row, col] = self.table.item(row, col).text()
        self.inventory.to_excel(INVENTORY_FILE, index=False)


# Running the app
if __name__ == "__main__":
    app = QApplication([])
    window = PokemonCardApp()
    window.show()
    app.exec_()
