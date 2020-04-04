import os
import time
import json
from threading import Thread

import tkinter as tk
from tkinter import BooleanVar, messagebox, filedialog, ttk

import docx
from docx.shared import Inches

from PIL import Image, ImageGrab, ImageTk
from pynput import keyboard
from plyer import notification

class SnipperTool():

    def __init__(self,
                 master,
                 title='Snipper Tool',
                 width=480,
                 height=350,
                 dict_image_format={'PNG':'.png', 'BMP':'.bmp', 'JPEG':'.jpeg'},
                 dict_image_editor_application={'MS Paint': 'mspaint'}):

        self.master = master
        self.title = title
        self.width = width
        self.height = height

        self.dict_image_format = dict_image_format
        self.dict_image_editor_application = dict_image_editor_application

        self.text_box_dir_path = None
        self.btn_browse = None
        self.btn_start = None
        self.btn_create_docx = None
        self.btn_exit = None
        self.chk_box_save_setings = BooleanVar()

        self.logo = None
        self.bg_image = None
        self.path_icon = None
        self.path_settings_file = None

        self.cmb_box_image_ext = None
        self.cmb_box_image_editor = None

        self.thread_start_process = None
        self.show_image_captured_notification = True # No usage at the moment

        self._draw_window()

    # region Drawing Window

    def _draw_window(self):

        self.master.title(self.title)
        self._set_file_path()

        self._set_window_size_and_position()
        self._draw_icon_and_logo()
        self._write_header_and_instructions()
        self._draw_text_box()
        self._draw_buttons()
        self._draw_settings()

    def _set_window_size_and_position(self):
        # get screen width and height
        width_screen = self.master.winfo_screenwidth()
        height_screen = self.master.winfo_screenheight()

        # calculate position (x and y coordinates) for the Tk root window
        x_window_pos = (width_screen/2) - (self.width/2)
        y_window_pos = (height_screen/2) - (self.height/2)

        # set the dimensions & position of window
        self.master.geometry('%dx%d+%d+%d' % (self.width, self.height, x_window_pos, y_window_pos))
        self.master.resizable(width=False, height=False)

    def _draw_icon_and_logo(self):

        if self.path_icon is not None and os.path.exists(self.path_icon):
            self.master.iconbitmap(self.path_icon)

            image_logo = (Image.open(self.path_icon)).resize((50, 50))
            self.logo = ImageTk.PhotoImage(image_logo)

            tk_label_logo = tk.Label(self.master, image=self.logo)
            tk_label_logo['image'] = self.logo
            tk_label_logo.place(x=48, y=6)

    def _create_label(self, label_text, x_text_pos, y_text_pos, font=("Helvetica", 8)):
        '''This will draw label on window'''
        label = tk.Label(self.master, text=label_text, font=font)
        label.place(x=x_text_pos, y=y_text_pos)

    def _write_header_and_instructions(self):
        # Header
        label_header = "Snipper Tool"
        self._create_label(label_header, 125, 13, font=("Arial", 30))

        # Instructions -----
        y_label = 225
        label_instruction_header = "Instructions :"
        self._create_label(label_instruction_header, 50, y_label)

        inst_1 = "1: Click on the 'Browse' button to select the folder."
        inst_2 = "2: Click on the ' Start ' button to start the process."
        inst_3 = "3: Press 'PrintScreen' key to capture the screenshot."
        inst_4 = "4: Press 'Insert' key to capture a screenshot & edit."
        inst_5 = "5: Click on 'Build Docx' button to create a document."
        inst_6 = "    With images available in selected folder."

        # label_instruction = inst_1 + '\n' + inst_2 + '\n' + inst_3 + '\n' + inst_4 + '\n' + inst_5
        # self.create_label(label_instruction, 50, 250)

        y_label = 245
        y_increment = 16
        self._create_label(inst_1, 50, y_label)
        self._create_label(inst_2, 50, y_label + y_increment * 1)
        self._create_label(inst_3, 50, y_label + y_increment * 2)
        self._create_label(inst_4, 50, y_label + y_increment * 3)
        self._create_label(inst_5, 50, y_label + y_increment * 4)
        self._create_label(inst_6, 50, y_label + y_increment * 5)

    def _draw_text_box(self):
        self.text_box_dir_path = tk.Entry(self.master, bd=2, width=46)
        self.text_box_dir_path.place(x=50, y=94)
        self.text_box_dir_path.delete(0, tk.END)

    def _create_button(self, button_text, x_pos, y_pos, button_width=10, button_click_command=None):
        button = tk.Button(self.master, text=button_text, command=button_click_command)
        button.config(width=button_width)
        button.place(x=x_pos, y=y_pos)

        return button

    def _draw_buttons(self):
        y_pos = 90
        self.btn_browse = self._create_button("Browse", 350, y_pos, 10, self._browse_dir)

        y_pos = y_pos + 85
        self.btn_start = self._create_button("Start", 50, y_pos, 17, self._start_capture_process)
        # self.btn_start.size(width = 130)
        # self.btn_new_folder = self._create_button("New Folder", 150, y, 10, self.create_new_dir)

        self.btn_create_docx = self._create_button("Create Docx", 200, y_pos,
                                                   17, self._create_document)

        self.btn_exit = self._create_button("Exit", 350, y_pos, 10, self.exit)

    def _draw_settings(self):
        y_pos = 135
        self._create_label('Type:  ', 50, y_pos, font=("Helvetica", 8))
        self.cmb_box_image_ext = ttk.Combobox(self.master,
                                              values=list(self.dict_image_format.keys()),
                                              state='readonly', font=("Helvetica", 8))
        self.cmb_box_image_ext.place(x=85, y=y_pos, width=95)

        self._create_label('Editor:', 198, y_pos, font=("Helvetica", 8))
        image_editors = list(self.dict_image_editor_application.keys())
        self.cmb_box_image_editor = ttk.Combobox(self.master,
                                                 values=image_editors,
                                                 state='readonly', font=("Helvetica", 8))
        self.cmb_box_image_editor.place(x=235, y=y_pos, width=94)

        y_pos = y_pos - 2
        chk_box_save = tk.Checkbutton(self.master, text='Save config',
                                      variable=self.chk_box_save_setings,
                                      onvalue=True, offvalue=False, command=self._save_setting)
        chk_box_save.place(x=345, y=y_pos)

        self._set_default_settings_value()

    # endregion Drawing Window


    # region Settings

    def _set_file_path(self):
        assets_path = os.path.join(os.getcwd(), 'assets')

        if not os.path.exists(assets_path):
            os.mkdir(assets_path)
        self.path_icon = os.path.join(assets_path, 'SnipperIcon.ico')

        # For executable or installer to have write access
        # Creating settings in user APPDATA folder
        path_settings = os.path.join(os.getenv('APPDATA'), self.title)
        if not os.path.exists(path_settings):
            os.mkdir(path_settings)
        self.path_settings_file = os.path.join(path_settings, 'settings.json')

    def _set_default_settings_value(self):
        self.cmb_box_image_ext.current(0)
        self.cmb_box_image_editor.current(0)
        self.chk_box_save_setings.set(False)

        self._load_settings()

        self.cmb_box_image_ext.bind("<<ComboboxSelected>>", self._save_setting)
        self.cmb_box_image_editor.bind("<<ComboboxSelected>>", self._save_setting)

    def _load_settings(self):
        if os.path.exists(self.path_settings_file):
            settings = None
            with open(self.path_settings_file, "r") as read_file:
                settings = json.load(read_file)

            index = list(self.dict_image_format.keys()).index(settings['image_type'])
            self.cmb_box_image_ext.current(index)

            index = list(self.dict_image_editor_application.keys()).index(settings['image_editor'])
            self.cmb_box_image_editor.current(index)

            self.chk_box_save_setings.set(settings['save_config'])

    def _save_setting(self, event=None):
        try:
            if self.chk_box_save_setings.get():
                settings = {'image_type': self.cmb_box_image_ext.get(),
                            'image_editor': self.cmb_box_image_editor.get(),
                            'save_config': self.chk_box_save_setings.get()}
                with open(self.path_settings_file, "w") as write_file:
                    json.dump(settings, write_file)
            else:
                if os.path.exists(self.path_settings_file):
                    os.remove(self.path_settings_file)
        except Exception as error:
            messagebox.showinfo(self.title, str(error))

    # endregion Settings


    def _browse_dir(self):
        dir_path = filedialog.askdirectory()
        self.text_box_dir_path.delete(0, tk.END)
        self.text_box_dir_path.insert(0, dir_path)

    # region Capture Process

    def _start_capture_process(self):
        if os.path.exists(self.text_box_dir_path.get()) and self.thread_start_process is None:
            self.thread_start_process = Thread(target=self._listen_key_events, daemon=True)
            self.thread_start_process.start()

            messagebox.showinfo(self.title, 'Process started.')
            self.master.iconify() # Minimises window tool
        elif not os.path.exists(self.text_box_dir_path.get()):
            messagebox.showinfo(self.title, "Please select correct path.")
        elif self.thread_start_process is not None:
            self._notify('Process already running.', 3, self.title, self.path_icon)

    def _listen_key_events(self):
        with keyboard.Listener(on_release=self._on_key_relase) as listener:
            listener.join()

    def _on_key_relase(self, key):
        if key == keyboard.Key.print_screen:
            # self._capture_screen(open_image=False)
            thread = Thread(target=self._capture_screen, args=(False, ), daemon=True)
            thread.start()
        elif key == keyboard.Key.insert:
            # self._capture_screen(open_image=True)
            thread = Thread(target=self._capture_screen, args=(True, ), daemon=True)
            thread.start()

    def _capture_screen(self, open_image=False, prefix_name='Screen_Shot_'):
        image_file_extn = self.dict_image_format[self.cmb_box_image_ext.get()]
        image_file_name = prefix_name + self.time_stamp() + image_file_extn
        image_file_path = os.path.join(self.text_box_dir_path.get(), image_file_name)

        ImageGrab.grab().save(image_file_path)

        if self.show_image_captured_notification:
            message = 'Image captured ' + image_file_name
            self._notify(message, 2, self.title, self.path_icon)

        if open_image:
            app_name = self.dict_image_editor_application[self.cmb_box_image_editor.get()]
            open_image_in_editor_cmd = app_name + ' ' + image_file_path
            os.system(open_image_in_editor_cmd)

    # endregion Capture Process


    # region Create Document

    def _create_document(self):
        """Creates word document wiht images present in selected directory"""

        if not os.path.exists(self.text_box_dir_path.get()):
            messagebox.showinfo(self.title, 'Please select correct path.')
            return

        list_path_images_in_dir = self._get_list_path_images()

        # If any images in selected directory
        if list_path_images_in_dir:
            doc = docx.Document()

            # Setting page layout of 1 inch margin from all sides
            for section in doc.sections:
                section.top_margin, section.bottom_margin = Inches(1), Inches(1)
                section.left_margin, section.right_margin = Inches(1), Inches(1)

            # Adding images to word document
            for index, file_path in enumerate(list_path_images_in_dir):
                doc.add_heading(str(index + 1), 4).style = 'List'
                doc.add_picture(file_path, width=Inches(6.5), height=Inches(3.65))
                doc.add_paragraph('') # For new line

            # Saving word document
            doc_name = 'Document_' + self.time_stamp() + '.docx'
            path_doc = os.path.abspath(os.path.join(self.text_box_dir_path.get(), doc_name))
            doc.save(path_doc)

            count_images = str(len(list_path_images_in_dir))
            msg = "Word document created with " + count_images + " images at " + path_doc + '.'
            messagebox.showinfo(self.title, msg)
        else:
            messagebox.showinfo(self.title, "No images present in selected directory.")

    def _get_list_path_images(self):
        list_path_images_in_dir = []
        image_extensions = tuple(list(self.dict_image_format.values()) + ['.jpg'])

        # Finding all images present in selected directory
        for file_name in os.listdir(self.text_box_dir_path.get()):
            if file_name.lower().endswith(image_extensions):
                path_image = os.path.abspath(os.path.join(self.text_box_dir_path.get(), file_name))
                list_path_images_in_dir.append(path_image)

        return list_path_images_in_dir

    # endregion Create Document

    def exit(self):
        """To close the application"""
        self.master.destroy()

    def time_stamp(self):
        """Returns time stamp with custom format as string"""
        return str(time.strftime("%Y-%m-%d-%H-%M-%S"))

    def _notify(self, message: str, timeout: int, title: str, app_icon: str) -> None:
        """
        Shows notification message

        Args:
            message (str):  the message to display
            timeout (int):  the time for which notification should be visible
            title (str):    title of notification
            app_icon (str): accessible path of icon file
        """
        notification.notify(title=title, message=message,
                            app_name=title, app_icon=app_icon,
                            timeout=timeout)

if __name__ == "__main__":
    MASTER = tk.Tk()
    SnipperTool(MASTER)
    MASTER.mainloop()
