import pyodbc
# import matplotlib.pyplot as plt
import os
from pathlib import Path
from PIL import Image as pil_image
import copy
import datetime
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
import yaml
import threading



class MainFrame(Frame):
    def __init__(self):
        super().__init__()
        self.stitching_running = False

        self.db_file_path_label = Label(self, text="Путь к файлу базы данных:")
        self.db_file_path_label.grid(row=0, columnspan=3)
        self.db_file_path = StringVar()
        self.db_file_path_field = Entry(self, textvariable=self.db_file_path, width=80)
        self.db_file_path_button = Button(self, command=self.browse_files, text='Выбрать')
        self.db_file_path_field.grid(row=1, column=1, columnspan=2)
        self.db_file_path_button.grid(row=1, column=0)

        self.images_folder_path_label = Label(self, text="\nПуть к одной папке с изображениями:")
        self.images_folder_path_label.grid(row=3, columnspan=3)
        self.images_folder_path = StringVar()
        self.images_folder_path_field = Entry(self, textvariable=self.images_folder_path, width=80)
        self.images_folder_path_button = Button(self, command=lambda: self.browse_directory('one_folder'), text='Выбрать')
        self.images_folder_path_field.grid(row=4, column=1, columnspan=2)
        self.images_folder_path_button.grid(row=4, column=0)

        self.dirs_folder_path_label = Label(self, text="\nПуть к нескольким папкам с изображениями:")
        self.dirs_folder_path_label.grid(row=5, columnspan=3)
        self.dirs_folder_path = StringVar()
        self.dirs_folder_path_field = Entry(self, textvariable=self.dirs_folder_path, width=80)
        self.dirs_folder_path_button = Button(self, command=lambda: self.browse_directory('multiple_folders'), text='Выбрать')
        self.dirs_folder_path_field.grid(row=6, column=1, columnspan=2)
        self.dirs_folder_path_button.grid(row=6, column=0)

        self.program_id_label = Label(self, text="\nID программы:")
        self.program_id_label.grid(row=7, column=0)
        self.program_id = IntVar()
        self.program_id_label_field = Entry(self, textvariable=self.program_id, width=6)
        self.program_id_label_field.grid(row=7, column=1, sticky='sw')

        self.run_stitching_button = Button(self, command=self.start_stitching_in_thread, text='Сшить изображения')
        self.run_stitching_button.grid(row=7, column=2, sticky='sw')

        self.progress_bar_label = Label(self, text='\n')
        self.progress_bar = Progressbar(self, orient=HORIZONTAL, length=100, mode='determinate')

        self.read_config_file()

        if Path.cwd().joinpath('result').exists() is False:
            Path.cwd().joinpath('result').mkdir()

    def browse_files(self):
        filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("DB file",
                                                          "*.mdb*"),
                                                         ("all files",
                                                          "*.*")))
        self.db_file_path.set(filename)

    def browse_directory(self, what):
        directory = filedialog.askdirectory()
        if what == 'one_folder':
            self.images_folder_path.set(directory)
        if what == 'multiple_folders':
            self.dirs_folder_path.set(directory)

    def browse_folders(self, multiple_directory):
        folders_list = []
        item_list = os.listdir(multiple_directory)
        for item in item_list:
            if Path.is_dir(Path(multiple_directory).joinpath(item)):
                folders_list.append(Path(multiple_directory).joinpath(item))
        return folders_list

    def image_stitching(self, **kwargs):
        """Initial things"""
        coord_catalog = []
        n_catalog = []
        x_catalog = []
        y_catalog = []
        db_file_path = self.db_file_path.get()
        db_driver = r'{Microsoft Access Driver (*.mdb, *.accdb)};'
        conn_str = f'DRIVER={db_driver}; DBQ={db_file_path};'

        '''Create DB connection'''
        conn = pyodbc.connect(conn_str)
        cur = conn.cursor()

        # for table_info in cur.tables(tableType='TABLE'): # get tables names in db
        #     print(table_info.table_name)

        # for row in cur.columns(table='Positions'): # get columns name in target table
        #     print(row.column_name)

        '''Making request to DB. Get program ID, positions ID, coordinates'''
        request = 'SELECT ID, ActualPosition, Axe1, Axe2, Axe3, Axe4 FROM Positions;'
        rows = cur.execute(request).fetchall()
        for row in rows:
            coord_catalog.append((row[0], row[1], -row[2], -row[3]))  # X and Y with minus because of mirroring

        '''Close DB connection'''
        cur.close()
        conn.close()

        '''Making lists with raw coordinates for chosen program'''
        program_id = self.program_id.get()
        for entry in coord_catalog:
            if entry[0] == program_id:
                n_catalog.append(entry[1])
                x_catalog.append(entry[2])
                y_catalog.append(entry[3])

        '''Calculate minimums'''
        x_min = copy.deepcopy(x_catalog)
        x_min = min(x_min)
        y_min = copy.deepcopy(y_catalog)
        y_min = min(y_min)

        '''Enter scale factor'''
        # TODO: Fixed scale values for X and Y axes, editable 'height' value!
        # scale_factor = 142.86 * 3.025  # Top fit
        scale_factor = 142.86 * 2.975  # Pipes fit  ******* This value can be edited *******
        print(f'Scale factor: {scale_factor}')

        '''Making normalized and scaled coordinates'''
        norm_x_catalog = [(x - x_min) / scale_factor for x in x_catalog]
        norm_y_catalog = [(y - y_min) / scale_factor / 2 for y in y_catalog]

        '''Making plot'''  # Currently turned off
        # plt.scatter(x_catalog, y_catalog)  # Get raw coordinates to plot
        # plt.scatter(norm_x_catalog, norm_y_catalog)  # Get normalized and scaled coordinates to plot
        # plt.ylabel('Y')
        # plt.xlabel('X')
        # plt.gca().set_aspect('equal', adjustable='box')  # This make plot scale factor equal for both axes
        # plt.show()

        '''Parsing images from directory'''
        images = {}
        cropping_size = 26  # Deleting left and top frame  ******* This value can be edited *******
        directory = Path(kwargs.get('directory'))
        files = os.listdir(directory)
        for (file, n_pos, x_pos, y_pos) in zip(files, n_catalog, norm_x_catalog, norm_y_catalog):
            if 'Pos0' in file:
                pos = file.find('Pos0')
                image_number = file[pos + 3: pos + 7]
                image = pil_image.open(directory.joinpath(file))
                # image = image.rotate(0.5)  # Turned off because quality decrease
                image = image.crop((cropping_size, cropping_size, 1024, 1024))  # Deleting borders
                image.n_pos = n_pos
                image.x_pos = x_pos
                image.y_pos = y_pos
                images.update({int(image_number): image})

        '''Making draw'''
        pillow_size = (int(max(norm_x_catalog)) + 1024 - cropping_size,
                       int(max(norm_y_catalog)) + 1024 - cropping_size + 500)  # 1024 - size of not cropped image
        pillow_image = pil_image.new('RGB', pillow_size)

        def y_inversing(norm_y):  # Convert normalized coordinates to pillow coordinate system (Y vice versa)
            inversed_y = max(norm_y_catalog) - norm_y
            return inversed_y

        '''**************'''
        '''Here supposed to make transparency for each image in size as pillow_image.
        Then blend each image to pillow_image. Maybe will do it later'''
        '''**************'''

        '''Pasting images to draw'''
        number_of_columns = 3  # ******* This value can be edited *******
        #  'number_of_columns' # TODO: supposed to be a function's return value in future
        column_counter = 0
        row_counter = 1
        row_shifting_px = 40  # ******* This value can be edited *******
        for image_num in images:
            column_counter += 1
            if column_counter > number_of_columns:
                row_counter += 1
                column_counter = 1
            image_shifting_px = 5  # If images have unfixed rotation ******* This value can be edited *******
            pillow_image.paste(images.get(image_num),
                               (int(images.get(image_num).x_pos) - image_shifting_px * (row_counter - 1),
                                int(y_inversing(images.get(image_num).y_pos)) +
                                image_shifting_px * column_counter +
                                row_counter * row_shifting_px))

        '''Saving image'''
        now_is = datetime.datetime.now().strftime('%Y_%m_%dT%H_%M_%S')
        last_folder_name = directory.parts[-1]
        image_file_type = '.png'  # ******* This value can be edited *******
        new_dir_name = global_path.joinpath("result")
        new_file_name = f'{last_folder_name}_result_{now_is}{image_file_type}'
        pillow_image.save(new_dir_name.joinpath(new_file_name))

    def image_stitching_from_multiple_dirs(self):
        folders_list = self.browse_folders(self.dirs_folder_path.get())
        for directory in folders_list:
            self.image_stitching(directory=directory)

    def image_stitching_from_one_dir(self):
        directory = self.images_folder_path.get()
        self.image_stitching(directory=directory)

    def read_config_file(self):
        try:
            print('Reading config file')
            with open('config.yaml', 'r') as file:
                data_dict = yaml.load(file, Loader=yaml.FullLoader)

        except:
            print('No config file found')
        try:
            db_dir = data_dict.get('db_dir')
            single_dir = data_dict.get('single_dir')
            multiple_dir = data_dict.get('multiple_dir')
            prog_num = data_dict.get('prog_num')
            self.db_file_path.set(db_dir)
            self.images_folder_path.set(single_dir)
            self.dirs_folder_path.set(multiple_dir)
            self.program_id.set(prog_num)
        except:
            print('Cant load values')

    def write_config_file(self):
        db_dir = self.db_file_path.get()
        single_dir = self.images_folder_path.get()
        multiple_dir = self.dirs_folder_path.get()
        prog_num = self.program_id.get()
        data_dict = {'db_dir': db_dir,
                     'single_dir': single_dir,
                     'multiple_dir': multiple_dir,
                     'prog_num': prog_num}
        with open('config.yaml', 'w') as file:
            yaml.dump(data_dict, file)

    def start_stitching(self):
        if len(self.dirs_folder_path.get()) != 0:
            self.image_stitching_from_multiple_dirs()
        if len(self.images_folder_path.get()) != 0:
            self.image_stitching_from_one_dir()
        self.write_config_file()
        self.stitching_running = False
        self.run_stitching_button.configure(state=NORMAL)

    def start_stitching_in_thread(self):
        self.stitching_running = True
        self.run_stitching_button.configure(state=DISABLED)
        self.start_stitching_thread = threading.Thread(target=self.start_stitching)
        self.start_stitching_thread.start()

    # TODO: Progress bar function

class MainApplication(Tk):
    def __init__(self):
        super().__init__()
        self.title('Image Stitching')

        self.frame = MainFrame()
        self.frame.grid()


app = MainApplication()
global_path = Path.cwd()
# app.geometry('700x700')  # wide x height
app.resizable(width=False, height=False)
app.mainloop()
