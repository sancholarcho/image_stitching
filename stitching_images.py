import pyodbc
# import matplotlib.pyplot as plt
import os
from PIL import Image as pil_image
import copy
import datetime
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog


class MainFrame(Frame):
    def __init__(self):
        super().__init__()

        self.db_file_path_label = Label(self, text="Путь к файлу базы данных:")
        self.db_file_path_label.grid(row=0, columnspan=3)
        self.db_file_path = StringVar()
        self.db_file_path_field = Entry(self, textvariable=self.db_file_path, width=80)
        self.db_file_path_button = Button(self, command=self.browse_files, text='Выбрать')
        self.db_file_path_field.grid(row=1, column=1, columnspan=2)
        self.db_file_path_button.grid(row=1, column=0)

        self.images_folder_path_label = Label(self, text="\nПуть к папке с изображениями:")
        self.images_folder_path_label.grid(row=3, columnspan=3)
        self.images_folder_path = StringVar()
        self.images_folder_path_field = Entry(self, textvariable=self.images_folder_path, width=80)
        self.images_folder_path_button = Button(self, command=self.browse_directory, text='Выбрать')
        self.images_folder_path_field.grid(row=4, column=1, columnspan=2)
        self.images_folder_path_button.grid(row=4, column=0)

        self.program_id_label = Label(self, text="\nID программы:")
        self.program_id_label.grid(row=5, column=0)
        self.program_id = IntVar()
        self.program_id_label_field = Entry(self, textvariable=self.program_id, width=6)
        self.program_id_label_field.grid(row=5, column=1, sticky='sw')

        self.run_stitching_label = Label(self, text='\n')
        self.run_stitching_label.grid(row=6, columnspan=2)
        self.run_stitching_button = Button(self, command=self.image_stitching, text='Сшить изображения')
        self.run_stitching_button.grid(row=5, column=2, sticky='sw')

    def browse_files(self):
        filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("DB file",
                                                          "*.mdb*"),
                                                         ("all files",
                                                          "*.*")))
        self.db_file_path.set(filename)

    def browse_directory(self):
        directory = filedialog.askdirectory()
        self.images_folder_path.set(directory)

    def image_stitching(self):
        """Initial things"""
        coord_catalog = []
        new_rows = []
        n_catalog = []
        x_catalog = []
        y_catalog = []
        # file_path = r'D:\access\Programs.mdb'  # ******* This value can be edited *******
        # file_path = input(r'Enter path to .mdb file (example: D:\access\Programs.mdb): ')
        file_path = self.db_file_path.get()
        db_driver = r'{Microsoft Access Driver (*.mdb, *.accdb)};'
        conn_str = f'DRIVER={db_driver}; DBQ={file_path};'

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
        # program_id = 5  # This is Program ID from DB ******* This value can be edited *******
        # program_id = input('Enter program ID (example: 5): ')
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
        # scale_factor = 142.86 * 3.025  # Top fit
        scale_factor = 142.86 * 2.975  # Pipes fit  ******* This value can be edited *******
        print(f'Scale factor: {scale_factor}')
        # Height from sensor to object instead of scale factor will be added in future

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
        # directory = r'D:/access/images'  # ******* This value can be edited *******
        # directory = input(r'Enter path to images (example: D:/access/images): ')
        directory = self.images_folder_path.get()
        files = os.listdir(directory)
        for (file, n_pos, x_pos, y_pos) in zip(files, n_catalog, norm_x_catalog, norm_y_catalog):
            if 'Pos0' in file:
                pos = file.find('Pos0')
                image_number = file[pos + 3:pos + 7]
                image = pil_image.open(directory + '/' + file)
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

        # alpha_image = pil_image.new('RGB', pillow_size, (255, 255, 255))

        def y_inversing(norm_y):  # Convert normalized coordinates to pillow coordinate system (Y vice versa)
            inversed_y = max(norm_y_catalog) - norm_y
            return inversed_y

        '''**************'''
        '''Here supposed to make transparency for each image in size as pillow_image.
        Then blend each image to pillow_image. Maybe will do it later'''
        '''**************'''

        '''Pasting images to draw'''
        number_of_columns = 3  # ******* This value can be edited *******
        #  'number_of_columns' supposed to be a function's return value in future. The solution will be added later
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
        pillow_image.save(f'result_{now_is}.png')


class MainApplication(Tk):
    def __init__(self):
        super().__init__()
        self.title('Image Stitching')

        self.frame = MainFrame()
        self.frame.grid()


app = MainApplication()
# app.geometry('700x700')  # wide x height
# app.resizable(width=True, height=False)
app.mainloop()
