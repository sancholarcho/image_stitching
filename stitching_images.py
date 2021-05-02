import pyodbc
import logging
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
        self.folders_number_total = 0
        self.stitching_running = False
        self.progress_total = 0

        self.db_file_path_label = Label(self, text="Путь к файлу базы данных:")
        self.db_file_path_label.grid(row=0, columnspan=4)
        self.db_file_path = StringVar()
        self.db_file_path_field = Entry(self, textvariable=self.db_file_path, width=80)
        self.db_file_path_button = Button(self, command=self.browse_files, text='Выбрать')
        self.db_file_path_field.grid(row=1, column=1, columnspan=3, padx=5)
        self.db_file_path_button.grid(row=1, column=0)

        self.images_folder_path_label = Label(self, text="\nПуть к одной папке с изображениями:")
        self.images_folder_path_label.grid(row=3, columnspan=4)
        self.images_folder_path = StringVar()
        self.images_folder_path_field = Entry(self, textvariable=self.images_folder_path, width=80)
        self.images_folder_path_button = Button(self, command=lambda: self.browse_directory(self.images_folder_path),
                                                text='Выбрать')
        self.images_folder_path_field.grid(row=4, column=1, columnspan=3, padx=5)
        self.images_folder_path_button.grid(row=4, column=0)

        self.dirs_folder_path_label = Label(self, text="\nПуть к нескольким папкам с изображениями:")
        self.dirs_folder_path_label.grid(row=5, columnspan=4)
        self.dirs_folder_path = StringVar()
        self.dirs_folder_path_field = Entry(self, textvariable=self.dirs_folder_path, width=80)
        self.dirs_folder_path_button = Button(self, command=lambda: self.browse_directory(self.dirs_folder_path),
                                              text='Выбрать')
        self.dirs_folder_path_field.grid(row=6, column=1, columnspan=3, padx=5)
        self.dirs_folder_path_button.grid(row=6, column=0)

        self.save_path_label = Label(self, text="\nПуть к папке сохранения:")
        self.save_path_label.grid(row=7, columnspan=4)
        self.save_path = StringVar()
        self.save_path_field = Entry(self, textvariable=self.save_path, width=80)
        self.save_path_button = Button(self, command=lambda: self.browse_directory(self.save_path),
                                       text='Выбрать')
        self.save_path_field.grid(row=8, column=1, columnspan=3, padx=5)
        self.save_path_button.grid(row=8, column=0)

        self.program_id_label = Label(self, text="\nID программы:")
        self.program_id_label.grid(row=9, column=0, sticky='se')
        self.program_id = IntVar()
        self.program_id_field = Entry(self, textvariable=self.program_id, width=6)
        self.program_id_field.grid(row=9, column=1, sticky='sw', padx=5)

        # self.height_label = Label(self, text="\nРасстояние от детектора до слоя изделия, мм:")
        self.height_label = Label(self, text="\nМасштабный коэффициент:")
        self.height_label.grid(row=10, column=0, sticky='se')
        self.height = DoubleVar()
        self.height_field = Entry(self, textvariable=self.height, width=20)
        self.height_field.grid(row=10, column=1, sticky='sw', padx=5)

        self.columns_label = Label(self, text="\nИзображений в ряду:")
        self.columns_label.grid(row=11, column=0, sticky='se')
        self.columns = IntVar()
        self.columns_field = Entry(self, textvariable=self.columns, width=6)
        self.columns_field.grid(row=11, column=1, sticky='sw', padx=5)

        self.row_shifting_label = Label(self, text="\nСмещение ряда, px:")
        self.row_shifting_label.grid(row=12, column=0, sticky='se')
        self.row_shifting = IntVar()
        self.row_shifting_field = Entry(self, textvariable=self.row_shifting, width=12)
        self.row_shifting_field.grid(row=12, column=1, sticky='sw', padx=5)

        self.run_stitching_button = Button(self, command=self.start_stitching_in_thread, text='Сшить изображения')
        self.run_stitching_button.grid(row=13, column=3)

        self.progress_bar_label = Label(self, text='\n')
        self.progress_bar_label.grid(row=13)
        self.progress_bar = Progressbar(self, orient=HORIZONTAL, length=400, mode='determinate')
        self.progress_bar.grid(row=13, column=0, columnspan=3, sticky='e', pady=20, padx=5)

        self.read_config_file()

    def browse_files(self):
        filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("DB file",
                                                          "*.mdb*"),
                                                         ("all files",
                                                          "*.*")))
        self.db_file_path.set(filename)

    def browse_directory(self, variable):
        directory = filedialog.askdirectory()
        variable.set(directory)

    def browse_folders(self, multiple_directory):
        folders_list = []
        item_list = os.listdir(multiple_directory)
        for item in item_list:
            if Path.is_dir(Path(multiple_directory).joinpath(item)):
                folders_list.append(Path(multiple_directory).joinpath(item))
        self.folders_number = len(folders_list)
        return folders_list

    def image_stitching(self, **kwargs):
        """Initial things"""
        logging.info('Start stitching')
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

        '''
        ^ Y
        |
        |
        -------> X
        '''

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
        # height = int(self.height.get())
        # focal_length = 850  # 2432 # TODO: May be need to clarify this value
        # step_x = 3564.55  # Calculated with ruler, x_value/mm
        # # step_y = 6800
        # step_y = 6168.44  # Calculated with ruler, y_value/mm
        # scale_factor = focal_length / height
        # scale_factor = 142.86 * 3.025  # Top fit
        # scale_factor = 142.86 * 2.975  # Pipes fit
        scale_factor = int(self.height.get())
        logging.info(f'Scale factor: {scale_factor:.2f}')

        '''Making normalized and scaled coordinates'''
        # norm_x_catalog = [((x - x_min) / step_x) * scale_factor for x in x_catalog]
        # norm_y_catalog = [((y - y_min) / step_y) * scale_factor for y in y_catalog]
        norm_x_catalog = [(x - x_min) / scale_factor for x in x_catalog]
        norm_y_catalog = [(y - y_min) / scale_factor / 2 for y in y_catalog]

        '''Making plot'''  # Currently turned off
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

        def y_inverting(norm_y):  # Convert normalized coordinates to pillow coordinate system (Y vice versa)
            inverted_y = max(norm_y_catalog) - norm_y
            return inverted_y

        '''
        X and Y pillow axes schema

        -------> X
        |
        |
        v Y
        '''

        '''**************'''
        '''Here supposed to make transparency for each image in size as pillow_image.
        Then blend each image to pillow_image. Maybe will do it later'''
        '''**************'''

        '''Pasting images to draw'''
        # number_of_columns = 3  # ******* This value can be edited *******
        number_of_columns = self.columns.get()
        image_shifting_px = -5  # If images have unfixed rotation ******* This value can be edited *******
        column_counter = 0
        row_counter = 1
        # row_shifting_px = 40  # TODO: check how this value influences with height change
        row_shifting_px = self.row_shifting.get()
        for image_num in images:
            column_counter += 1
            if column_counter > number_of_columns:
                row_counter += 1
                column_counter = 1
            logging.info(f'X={images.get(image_num).x_pos:.2f}, Y={y_inverting(images.get(image_num).y_pos):.2f}')
            dx = image_shifting_px * (row_counter - 1)
            dy = - image_shifting_px * (column_counter - 1) + row_counter * row_shifting_px
            logging.info(f'dX={dx}, dY={dy}')
            x_now = round(images.get(image_num).x_pos) + dx
            y_now = round(y_inverting(images.get(image_num).y_pos)) + dy
            logging.info(f'X={x_now}, Y={y_now}')
            logging.info('-----------')
            pillow_image.paste(images.get(image_num), (x_now, y_now))
        '''Saving image'''
        now_is = datetime.datetime.now().strftime('%Y_%m_%dT%H_%M_%S')
        last_folder_name = directory.parts[-1]
        image_file_type = '.png'  # ******* This value can be edited *******
        new_dir_name = Path(self.save_path.get()).joinpath('result')
        new_file_name = f'{last_folder_name}_result_{now_is}{image_file_type}'
        if new_dir_name.exists() is False:
            new_dir_name.mkdir()
        pillow_image.save(new_dir_name.joinpath(new_file_name))

        '''Progressbar'''
        self.progress = 100 / self.folders_number_total
        self.progress_total += self.progress
        self.progress_bar['value'] = self.progress_total
        self.update_idletasks()

        logging.info('Image stitched')

    def image_stitching_from_multiple_dirs(self):
        folders_list = self.browse_folders(self.dirs_folder_path.get())
        self.folders_number_total += len(folders_list)
        for directory in folders_list:
            self.image_stitching(directory=directory)

    def image_stitching_from_one_dir(self):
        directory = self.images_folder_path.get()
        self.folders_number_total += 1
        self.image_stitching(directory=directory)

    def read_config_file(self):
        try:
            logging.info('Reading config file')
            with open('config.yaml', 'r') as file:
                data_dict = yaml.load(file, Loader=yaml.FullLoader)

        except FileNotFoundError:
            logging.warning('No config file found')
        try:
            db_dir = data_dict.get('db_dir')
            single_dir = data_dict.get('single_dir')
            multiple_dir = data_dict.get('multiple_dir')
            program_num = data_dict.get('program_num')
            save_dir = data_dict.get('save_dir')
            height = data_dict.get('height')
            columns = data_dict.get('columns')
            row_shifting = data_dict.get('row_shifting')
            self.db_file_path.set(db_dir)
            self.images_folder_path.set(single_dir)
            self.dirs_folder_path.set(multiple_dir)
            self.program_id.set(program_num)
            self.save_path.set(save_dir)
            self.height.set(height)
            self.columns.set(columns)
            self.row_shifting.set(row_shifting)
        except (UnboundLocalError, KeyError):
            self.db_file_path.set(None)
            self.images_folder_path.set(None)
            self.dirs_folder_path.set(None)
            self.program_id.set(None)
            self.save_path.set(None)
            self.height.set(None)
            self.columns.set(None)
            self.row_shifting.set(None)
            logging.info('Cant load values, setting None')

    def write_config_file(self):
        db_dir = self.db_file_path.get()
        single_dir = self.images_folder_path.get()
        multiple_dir = self.dirs_folder_path.get()
        program_num = self.program_id.get()
        save_dir = self.save_path.get()
        height = self.height.get()
        columns = self.columns.get()
        row_shifting = self.row_shifting.get()
        data_dict = {'db_dir': db_dir,
                     'single_dir': single_dir,
                     'multiple_dir': multiple_dir,
                     'program_num': program_num,
                     'save_dir': save_dir,
                     'height': height,
                     'columns': columns,
                     'row_shifting': row_shifting}
        with open('config.yaml', 'w') as file:
            yaml.dump(data_dict, file)

    def start_stitching(self):
        self.progress_total = 0
        if len(self.dirs_folder_path.get()) != 0:
            self.image_stitching_from_multiple_dirs()
        if len(self.images_folder_path.get()) != 0:
            self.image_stitching_from_one_dir()
        self.write_config_file()
        self.stitching_running = False
        self.run_stitching_button.configure(state=NORMAL)
        self.folders_number_total = 0

    def start_stitching_in_thread(self):
        self.stitching_running = True
        self.run_stitching_button.configure(state=DISABLED)
        self.start_stitching_thread = threading.Thread(target=self.start_stitching)
        self.start_stitching_thread.start()


class MainApplication(Tk):

    def __init__(self):
        super().__init__()
        self.title('Image Stitching')

        self.frame = MainFrame()
        self.frame.grid()


logging.basicConfig(level=logging.DEBUG)
app = MainApplication()
app.resizable(width=False, height=False)
app.mainloop()
