import random
import re

from PyQt6.QtGui import QStandardItemModel, QStandardItem
from PyQt6.QtWidgets import QDialog, QAbstractItemView

from linvam.ui_soundactioneditwnd import Ui_SoundSelect
from linvam.util import get_voice_packs_folder_path, Command


class SoundActionEditWnd(QDialog):
    def __init__(self, p_sounds, p_sound_action=None, p_parent=None):
        super().__init__(p_parent)
        self.ui = Ui_SoundSelect()
        self.ui.setupUi(self)

        if p_sounds is None:
            return

        self.p_sounds = p_sounds
        self.selected_voice_pack = None
        self.selected_category = None
        self.selected_files = []  # Changed to list to support multiple files
        self.m_sound_action = {}

        self.ui.buttonOkay.clicked.connect(self.slot_ok)
        self.ui.buttonCancel.clicked.connect(super().reject)
        self.ui.buttonPlaySound.clicked.connect(self.play_sound)
        self.ui.buttonStopSound.clicked.connect(self.stop_sound)
        self.ui.buttonPlaySound.setEnabled(False)
        self.ui.buttonStopSound.setEnabled(False)
        self.ui.buttonOkay.setEnabled(False)

        # restore stuff when editing
        if p_sound_action is not None:
            self.selected_voice_pack = p_sound_action['pack']
            self.selected_category = p_sound_action['cat']
            # Support both old single-file format and new multi-file format
            if 'files' in p_sound_action:
                self.selected_files = p_sound_action['files'][:]
            elif 'file' in p_sound_action:
                self.selected_files = [p_sound_action['file']]
            self.ui.buttonOkay.setEnabled(True)

        self.list_voice_packs_model = QStandardItemModel()
        self.ui.listVoicepacks.setModel(self.list_voice_packs_model)
        self.ui.listVoicepacks.clicked.connect(self.on_voice_pack_select)

        self.list_categories_model = QStandardItemModel()
        self.ui.listCategories.setModel(self.list_categories_model)
        self.ui.listCategories.clicked.connect(self.on_category_select)

        self.list_files_model = QStandardItemModel()
        self.ui.listFiles.setModel(self.list_files_model)
        # Enable multi-selection for files
        self.ui.listFiles.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.ui.listFiles.clicked.connect(self.on_file_select)
        self.ui.listFiles.doubleClicked.connect(self.select_and_play)

        s = sorted(p_sounds.m_sounds)
        for v in s:
            item = QStandardItem(v)
            self.list_voice_packs_model.appendRow(item)

        self.ui.filterCategories.textChanged.connect(self.populate_categories)
        self.ui.filterFiles.textChanged.connect(self.populate_files)

        self.populate_categories(False)
        self.populate_files(False)
        self.select_old_entries()

    def slot_ok(self):
        self.m_sound_action = {
            'name': Command.PLAY_SOUND,
            'pack': self.selected_voice_pack,
            'cat': self.selected_category,
            'files': self.selected_files  # Save as array of files
        }
        super().accept()

    def slot_cancel(self):
        super().reject()

    def on_voice_pack_select(self):
        index = self.ui.listVoicepacks.currentIndex()
        item_text = index.data()
        self.selected_voice_pack = item_text
        self.populate_categories()
        self.ui.buttonOkay.setEnabled(False)
        self.ui.buttonPlaySound.setEnabled(False)

    def on_category_select(self):
        index = self.ui.listCategories.currentIndex()
        item_text = index.data()
        self.selected_category = item_text
        self.populate_files()
        self.ui.buttonOkay.setEnabled(False)
        self.ui.buttonPlaySound.setEnabled(False)

    def on_file_select(self):
        # Get all selected files
        selected_indexes = self.ui.listFiles.selectedIndexes()
        self.selected_files = [index.data() for index in selected_indexes]

        # Enable buttons if at least one file is selected
        if self.selected_files:
            self.ui.buttonOkay.setEnabled(True)
            self.ui.buttonPlaySound.setEnabled(True)
        else:
            self.ui.buttonOkay.setEnabled(False)
            self.ui.buttonPlaySound.setEnabled(False)

    def select_and_play(self):
        self.on_file_select()
        self.play_sound()

    def populate_categories(self, reset=True):
        if self.selected_voice_pack is None:
            return

        if reset:
            self.list_categories_model.removeRows(0, self.list_categories_model.rowCount())
            self.list_files_model.removeRows(0, self.list_files_model.rowCount())
            self.selected_category = None
            self.selected_files = []

        filter_categories = self.ui.filterCategories.toPlainText()
        if len(filter_categories) == 0:
            filter_categories = None

        s = sorted(self.p_sounds.m_sounds[self.selected_voice_pack])
        for v in s:
            if filter_categories is not None:
                if not re.search(filter_categories, v, re.IGNORECASE):
                    continue
            item = QStandardItem(v)
            self.list_categories_model.appendRow(item)

    def populate_files(self, reset=True):
        if self.selected_voice_pack is None or self.selected_category is None:
            return

        if reset:
            self.list_files_model.removeRows(0, self.list_files_model.rowCount())
            self.selected_files = []

        filter_files = self.ui.filterFiles.toPlainText()
        if len(filter_files) == 0:
            filter_files = None

        s = sorted(self.p_sounds.m_sounds[self.selected_voice_pack][self.selected_category])
        for v in s:
            if filter_files is not None:
                if not re.search(filter_files, v, re.IGNORECASE):
                    continue
            item = QStandardItem(v)
            self.list_files_model.appendRow(item)

    def play_sound(self):
        if not self.selected_files:
            return

        # If multiple files selected, randomly choose one to preview
        selected_file = random.choice(self.selected_files)

        sound_file = (get_voice_packs_folder_path() + self.selected_voice_pack + '/' + self.selected_category + '/'
                      + selected_file)
        self.p_sounds.play(sound_file)
        self.ui.buttonStopSound.setEnabled(True)

    def stop_sound(self):
        self.p_sounds.stop()

    def select_old_entries(self):
        # when editing, select old entries
        if self.selected_voice_pack is not None:
            item = self.list_voice_packs_model.findItems(self.selected_voice_pack)
            if len(item) > 0:
                index = self.list_voice_packs_model.indexFromItem(item[0])
                self.ui.listVoicepacks.setCurrentIndex(index)

        if self.selected_category is not None:
            item = self.list_categories_model.findItems(self.selected_category)
            if len(item) > 0:
                index = self.list_categories_model.indexFromItem(item[0])
                self.ui.listCategories.setCurrentIndex(index)

        # Select multiple files if they were previously selected
        if self.selected_files:
            for file_name in self.selected_files:
                items = self.list_files_model.findItems(file_name)
                if len(items) > 0:
                    index = self.list_files_model.indexFromItem(items[0])
                    self.ui.listFiles.selectionModel().select(
                        index,
                        self.ui.listFiles.selectionModel().SelectionFlag.Select
                    )
            self.ui.buttonPlaySound.setEnabled(True)
