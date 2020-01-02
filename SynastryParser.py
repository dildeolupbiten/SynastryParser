#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__version__ = "1.0.0"

import os
import sys
import ssl
import time
import logging
import platform
import threading
import subprocess
import webbrowser
import tkinter as tk
import urllib.request
import tkinter.messagebox as msgbox

from tkinter import filedialog
from tkinter.ttk import Progressbar
from datetime import datetime as dt

logging.basicConfig(
    filename="output.log",
    format="- %(levelname)s - %(asctime)s - %(message)s",
    level=logging.DEBUG,
    datefmt="%d.%m.%Y %H:%M:%S"
)
logging.info("Session started.")

whl_path = os.path.join(os.getcwd(), "Eph", "Whl")


def select_module(module_name, module_files):
    if os.name == "posix":
        os.system(f"pip3 install {module_name}")
    elif os.name == "nt":
        if sys.version_info.minor == 6:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(whl_path, module_files[0])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(whl_path, module_files[1])
                os.system(f"pip3 install {new_path}")
        elif sys.version_info.minor == 7:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(whl_path, module_files[2])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(whl_path, module_files[3])
                os.system(f"pip3 install {new_path}")
        elif sys.version_info.minor == 8:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(whl_path, module_files[4])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(whl_path, module_files[5])
                os.system(f"pip3 install {new_path}")

try:
    import numpy as np
except ModuleNotFoundError:
    os.system("pip3 install numpy")
    import numpy as np
try:
    import xlwt
except ModuleNotFoundError:
    os.system("pip3 install xlwt")
    import xlwt
try:
    import shapely
except ModuleNotFoundError:            
    select_module(
        "shapely", 
        [i for i in os.listdir(whl_path) if "Shapely" in i]
    )      
try:
    import swisseph as swe
except ModuleNotFoundError:
    select_module(
        "pyswisseph",
        [i for i in os.listdir(whl_path) if "pyswisseph" in i]
    )
    import swisseph as swe    
    
swe.set_ephe_path(os.path.join(os.getcwd(), "Eph"))

CONJUNCTION = 10
SEMI_SEXTILE = 3
SEMI_SQUARE = 3
SEXTILE = 6
QUINTILE = 2
SQUARE = 10
TRINE = 10
SESQUIQUADRATE = 3
BIQUINTILE = 2
QUINCUNX = 3
OPPOSITE = 10

SIGNS = [
    "Aries",
    "Taurus",
    "Gemini",
    "Cancer",
    "Leo",
    "Virgo",
    "Libra",
    "Scorpio",
    "Sagittarius",
    "Capricorn",
    "Aquarius",
    "Pisces"
]

ANTISCIA = {
    SIGNS[0]: SIGNS[5],
    SIGNS[1]: SIGNS[4],
    SIGNS[2]: SIGNS[3],
    SIGNS[3]: SIGNS[2],
    SIGNS[4]: SIGNS[1],
    SIGNS[5]: SIGNS[0],
    SIGNS[6]: SIGNS[11],
    SIGNS[7]: SIGNS[10],
    SIGNS[8]: SIGNS[9],
    SIGNS[9]: SIGNS[8],
    SIGNS[10]: SIGNS[7],
    SIGNS[11]: SIGNS[6]
}

CONTRA_ANTISCIA = {
    SIGNS[0]: SIGNS[11],
    SIGNS[1]: SIGNS[10],
    SIGNS[2]: SIGNS[9],
    SIGNS[3]: SIGNS[8],
    SIGNS[4]: SIGNS[7],
    SIGNS[5]: SIGNS[6],
    SIGNS[6]: SIGNS[5],
    SIGNS[7]: SIGNS[4],
    SIGNS[8]: SIGNS[3],
    SIGNS[9]: SIGNS[2],
    SIGNS[10]: SIGNS[1],
    SIGNS[11]: SIGNS[0]
}

OBJECTS = [
    "Sun",
    "Moon",
    "Mercury",
    "Venus",
    "Mars",
    "Jupiter",
    "Saturn",
    "Uranus",
    "Pluto",
    "Mean",
    "True",
    "Chiron",
    "Juno",
    "Asc.",
    "MC"
]

ASPECTS = [
    "conjunction",
    "semi_sextile",
    "semi_square",
    "sextile",
    "quintile",
    "square",
    "trine",
    "sesquiquadrate",
    "biquintile",
    "quincunx",
    "opposite",
    "null"
]

ANGLE = {
    ASPECTS[0]: 0,
    ASPECTS[1]: 30,
    ASPECTS[2]: 45,
    ASPECTS[3]: 60,
    ASPECTS[4]: 72,
    ASPECTS[5]: 90,
    ASPECTS[6]: 120,
    ASPECTS[7]: 135,
    ASPECTS[8]: 144,
    ASPECTS[9]: 150,
    ASPECTS[10]: 180,
}

SST = {
    SIGNS[i]: 30 * i
    for i in range(len(SIGNS))
}
    

class Chart:
    __PLANET_DICT = {
        "Sun": swe.SUN,
        "Moon": swe.MOON,
        "Mercury": swe.MERCURY,
        "Venus": swe.VENUS,
        "Mars": swe.MARS,
        "Jupiter": swe.JUPITER,
        "Saturn": swe.SATURN,
        "Uranus": swe.URANUS,
        "Neptune": swe.NEPTUNE,
        "Pluto": swe.PLUTO,
        "North Node": swe.TRUE_NODE,
        "Chiron": swe.CHIRON,
        "Juno": swe.JUNO
    }

    def __init__(
            self, 
            jd: float = 0,
            lat: float = 0,
            lon: float = 0,
            hsys: str = ""
    ):
        self.jd = jd
        self.lat = lat
        self.lon = lon
        self.hsys = hsys

    @classmethod
    def convert_angle(cls, angle: float = 0):
        for i in range(12):
            if i * 30 <= angle < (i + 1) * 30:
                return angle - (30 * i), SIGNS[i]

    @classmethod
    def reverse_convert_angle(cls, degree: float = 0, sign: str = ""):
        return degree + 30 * SIGNS.index(sign)

    def planet_pos(self, planet: int = 0):
        calc = self.convert_angle(angle=swe.calc_ut(self.jd, planet)[0])
        return calc[1], self.reverse_convert_angle(calc[0], calc[1])

    def house_pos(self):
        house = []
        asc = 0
        angle = []
        for i, j in enumerate(swe.houses(
                self.jd, self.lat, self.lon,
                bytes(self.hsys.encode("utf-8")))[0]):
            if i == 0:
                asc += j
            angle.append(j)
            house.append((
                f"{i + 1}",
                j,
                f"{self.convert_angle(j)[1]}"))
        return house, asc, angle

    def patterns(self):
        count = 0
        planet_positions = []
        house_positions = []
        for key, value in self.__PLANET_DICT.items():
            planet = self.planet_pos(planet=value)
            planet_info = [
                key,
                planet[0],
                planet[1]
            ]
            planet_positions.append(planet_info)
            if count < 12:
                house = [
                    int(self.house_pos()[0][count][0]),
                    self.house_pos()[0][count][-1],
                    float(self.house_pos()[0][count][1]),                  
                ]
                house_positions.append(house)
            else:
                pass
            count += 1
        planet_positions.extend(
            [
                ["Asc.", *house_positions[0][1:]], 
                ["MC", *house_positions[9][1:]]
            ]
        )        
        return planet_positions
           
       
def find_aspect(natal1: list = [], natal2: list = []):
    result = {}
    for i in natal1:
        result[i[0]] = {}
        for j in natal2:
            result[i[0]][j[0]] = abs(i[2] - j[2])
    return result


def parse_aspects(synastry: dict = {}, first: int = 0, aspect: int = 0):
    result = {}
    for i, j in synastry.items():
        result[i] = []
        for k, m in j.items():
            if first == 0:
                if (first < m < aspect) or \
                        (360 - first - aspect < m < 360 - first):
                    result[i].append(1)
                else:
                    result[i].append(0)
            elif first == 180:
                if first - aspect < m < first + aspect:
                    result[i].append(1)
                else:
                    result[i].append(0)
            else:
                if (first - aspect < m < first + aspect) or \
                        (360 - first - aspect < m < 360 - first + aspect):
                    result[i].append(1)
                else:
                    result[i].append(0)
    return result


def search_aspect(x_: int = 0, y_: int = 0, orb: int = 0, aspect: int = 0):
    result = abs(x_ - y_)
    if aspect == 0:
        if (aspect < result < orb) or \
                (360 - orb < result < 360):
            return 1
        else:
            return 0
    elif aspect == 180:
        if aspect - orb < result < aspect + orb:
            return 1
        else:
            return 0
    else:
        if (aspect - orb < result < aspect + orb) or \
                (360 - aspect - orb < result < 360 - aspect + orb):
            return 1
        else:
            return 0


def create_table(_x: int = [], _y: int = [], obj: str = "", 
        aspect: int = 0, orb: int = 0
):
    result = []
    for k in range(len(OBJECTS)):
        arr = [[0 for _ in range(12)] for __ in range(12)]
        i = OBJECTS.index(obj)
        arr[SIGNS.index(_y[k][1])][SIGNS.index(_x[i][1])] = \
            search_aspect(
                x_=_x[i][2], 
                y_=_y[k][2], 
                orb=orb, 
                aspect=aspect
            )
        match = {
            SIGNS[j]: arr[j]
            for j in range(len(SIGNS))
        }
        result.append(match)
    return result


def change_mode(info: list = [], mode: dict = {}):
    for i, j in enumerate(info):
        info[i] = [info[i][0]] + \
                  [mode[info[i][1]]] + \
                  [
                      SST[mode[info[i][1]]] +
                      30 - info[i][2] % 30
                  ]
    return info


def change_values(selection: str= "", info: list = []):
    if selection == "Antiscia":
        info = change_mode(info=info, mode=ANTISCIA)
    elif selection == "Contra-Antiscia":
        info = change_mode(info=info, mode=CONTRA_ANTISCIA)
    return info


def activate_selections(selected: list = [], modes: list = [], 
        male: list = [], female: list = []
):
    global x, y
    x, y = male, female
    x = change_values(selection=modes[0], info=x)
    y = change_values(selection=modes[1], info=y)
    aspect_commands = {
        ASPECTS[0]: "parse_aspects("
                    "find_aspect(x, y), 0, CONJUNCTION)",
        ASPECTS[1]: "parse_aspects("
                    "find_aspect(x, y), 30, SEMI_SEXTILE)",
        ASPECTS[2]: "parse_aspects("
                    "find_aspect(x, y), 45, SEMI_SQUARE)",
        ASPECTS[3]: "parse_aspects("
                    "find_aspect(x, y), 60, SEXTILE)",
        ASPECTS[4]: "parse_aspects("
                    "find_aspect(x, y), 72, QUINTILE)",
        ASPECTS[5]: "parse_aspects("
                    "find_aspect(x, y), 90, SQUARE)",
        ASPECTS[6]: "parse_aspects("
                    "find_aspect(x, y), 120, TRINE)",
        ASPECTS[7]: "parse_aspects("
                    "find_aspect(x, y), 135, SESQUIQUADRATE)",
        ASPECTS[8]: "parse_aspects("
                    "find_aspect(x, y), 144, BIQUINTILE)",
        ASPECTS[9]: "parse_aspects("
                    "find_aspect(x, y), 150, QUINCUNX)",
        ASPECTS[10]: "parse_aspects("
                     "find_aspect(x, y), 180,OPPOSITE)"
    }
    result = []
    signs = []
    for k, v in aspect_commands.items():
        if k in selected:
            result += eval(v),
            for i in OBJECTS:
                if i in selected:
                    code = f"create_table(_x=x, _y=y, obj='{i}', "\
                           f"aspect=ANGLE['{k}'], orb={k.upper()})"
                    signs += eval(code),
    return result, signs


def sum_arrays(*args):
    result = []
    for i in range(len(args[0])):
        count = 0
        for j in args:
            count += j[i]
        result.append(count)
    return result


def count_aspects_1(file_list: list = []):
    result = ()
    arg = ()
    for i in file_list:
        if file_list.index(i) != len(file_list) - 1:
            if len(list(result)) == 0:
                result = file_list[file_list.index(i)]
            else:
                result = [i for i in arg]
                arg = ()
            for j, k in zip(
                    result, 
                    file_list[file_list.index(i) + 1]
            ):
                aspect = {}
                for (m, n), (o, p) in zip(j.items(), k.items()):
                    aspect[m] = sum_arrays(n, p)
                arg += aspect,
    return result
    

def count_aspects_2(split_list: list = [], selected: list = []):
    final = []
    count = 0
    for _i, _j in enumerate([s for s in selected if s in ASPECTS]):
        for i_, j_ in enumerate([o for o in selected if o in OBJECTS]):
            result = [
                {
                    j: np.array([0 for __ in range(12)]) for j in SIGNS
                } for _ in OBJECTS
            ]
            for i in split_list:
                for ind, (j, k) in enumerate(zip(i[count], result)):
                    for m, n in j.items():
                        result[ind][m] += np.array(i[count][ind][m])
            for i in range(len(result)):
                for j, m in result[i].items():
                    result[i][j] = [int(i) for i in m]
            final.append(result)
            count += 1
    return final
    
    
def read_files(file: str = "", selected: list = [], modes: list = [], 
        master=None
):
    file_list = []
    split_list = []
    if not selected or \
            all(i not in selected for i in ASPECTS) or \
            len(selected) == 1 and "all" in selected:
        msgbox.showinfo(
            message="No aspect selected."
        )
        return None
    else:
        try:
            file = open(file, "r+")
        except TypeError:
            return None
        readlines = file.readlines()
        n = len(readlines)
        logging.info(f"Number of charts: {n}")
        logging.info(f"Selected: {', '.join(selected)}")
        logging.info(f"Mode: {', '.join(modes)}")
        logging.info("Calculation started.")
        r = 0
        t = time.time()
        l = lambda: time.time()
        s = lambda: f"{int(100 * r / n)} %, "\
            f"{round((int(n / (r / (l() - t))) - int(l() - t)) / 60)} "\
            f"minutes remaining."
        pbar = Progressbar(
            master=master,
            orient="horizontal",
            length=200,
            mode="determinate"
        )
        pstring = tk.StringVar()
        plabel = tk.Label(master=master, textvariable=pstring)
        pbar.pack(side="left")
        plabel.pack(side="left")
        for i in range(0, n, 2):
            male = [float(j) for j in readlines[i].strip().split(",")]
            female = [float(j) for j in readlines[i + 1].strip().split(",")]
            splitted = activate_selections(
                selected=selected,
                male=Chart(*male, "P").patterns(),
                female=Chart(*female, "P").patterns(),
                modes=modes
            )
            split_list.append(
                splitted[1]
            )
            file_list.append(
                splitted[0]
            )
            r += 2
            if r != n:
                pbar["value"] = r
                pbar["maximum"] = n
                pstring.set(s())
            else:
                pbar["value"] = n
                pstring.set("100 %, 0 minutes remaining.")
                logging.info("Calculation finished.")
                master.update()
                msgbox.showinfo(message="Calculation finished.")
                pbar.destroy()
                plabel.destroy()
        return count_aspects_1(file_list), \
            count_aspects_2(split_list, selected)


class Spreadsheet(xlwt.Workbook):
    size = 3
    obj = None

    def __init__(self):
        xlwt.Workbook.__init__(self)
        self.sheet = self.add_sheet("Sheet1")
        self.style = xlwt.XFStyle()
        self.alignment = xlwt.Alignment()
        self.alignment.horz = xlwt.Alignment.HORZ_CENTER
        self.alignment.vert = xlwt.Alignment.VERT_CENTER
        self.style.alignment = self.alignment

    @staticmethod
    def font(name: str = "Arial", bold: bool = False):
        f = xlwt.Font()
        f.name = name
        f.bold = bold
        return f

    def write(self, table: dict = {}, aspect: str = ""):
        temporary = {}
        transpose = np.transpose([i for i in table.values()])
        arr = [[int(j) for j in list(i)] for i in transpose]
        for i, j in enumerate(table):
            temporary[j] = arr[i]
        if aspect == "all":
            label = f"{aspect.upper()}"
        else:
            label = f"{aspect.upper()} "\
                    f"(Orb Factor: +- {eval(aspect.upper())})"
        self.sheet.write_merge(
            r1=0 + self.size,
            r2=0 + self.size,
            c1=0,
            c2=len(temporary.keys()),
            label=label,
            style=self.style
        )
        _col = 0
        if aspect in OBJECTS:
            _col += 1
            self.sheet.write_merge(
                r1=2 + self.size,
                r2=1 + self.size + len(temporary.keys()),
                c1=0,
                c2=0,
                label=next(self.obj),
                style=self.style
            )
        row = 2 + self.size
        for index, (keys, values) in enumerate(temporary.items()):
            self.sheet.write(
                r=1 + self.size,
                c=index + 1 + _col,
                label=keys,
                style=self.style
            )
            self.sheet.write(
                r=index + 2 + self.size,
                c=0 + _col,
                label=keys,
                style=self.style
            )
            col = 1 + _col
            for value in values:
                self.sheet.write(
                    r=row,
                    c=col,
                    label=value,
                    style=self.style
                )
                col += 1
            row += 1
        self.size += len(table.keys()) + 3

    def create_tables(self, file: str = "", selected: list = [], 
            modes: list = [], master=None
    ):
        read = read_files(file, selected, modes, master)
        self.style.font = self.font(bold=True)
        self.sheet.write_merge(
            r1=0, c1=0, r2=1, c2=0, label="Mode", style=self.style
        )
        self.sheet.write(r=0, c=1, label="1. Person", style=self.style)
        self.sheet.write(r=0, c=2, label="2. Person", style=self.style)
        self.style.font = self.font(bold=False)
        self.sheet.write(r=1, c=1, label=modes[0], style=self.style)
        self.sheet.write(r=1, c=2, label=modes[1], style=self.style)
        try:
            tables = read[0]
            signs = read[1]
            if tables:
                if "all" in selected:
                    result = {}
                    for i in tables[0]:
                        result[i] = []
                        for j in tables:
                            for k, v in j.items():
                                if k == i:
                                    result[i].append(v)
                        result[i] = sum_arrays(*result[i])
                    tables += result,
                    self.write(
                        tables[-1],
                        "all",
                    )
                count = 0
                for index, aspect in enumerate(ASPECTS):
                    if aspect in selected:
                        self.write(
                            tables[selected.index(aspect)],
                            aspect,
                        )
                        for i in OBJECTS:
                            if i in selected:
                                for j in OBJECTS:
                                    self.sheet.write_merge(
                                        r1=self.size + 2,
                                        r2=self.size + 13,
                                        c1=0,
                                        c2=0,
                                        label=j,
                                        style=self.style
                                    )
                                    self.sheet.write_merge(
                                        r1=self.size,
                                        r2=self.size,
                                        c1=2,
                                        c2=13,
                                        label=i,
                                        style=self.style
                                    )
                                    for k in SIGNS:
                                        self.sheet.write(
                                            r=self.size + 1,
                                            c=2 + SIGNS.index(k),
                                            label=k,
                                            style=self.style
                                        )
                                        self.sheet.write(
                                            r=self.size + 2 + SIGNS.index(k),
                                            c=1,
                                            label=k,
                                            style=self.style
                                        )
                                    row = 0
                                    for key, value in signs[count][
                                            OBJECTS.index(j)
                                    ].items():
                                        for ind, val in enumerate(value):
                                            self.sheet.write(
                                                r=self.size + 2 + row,
                                                c=2 + ind,
                                                label=val,
                                                style=self.style
                                            )
                                        row += 1
                                    self.size += 15
                                count += 1
                self.save("synastry.xlsx")                
            else:
                pass
        except TypeError:
            pass


class App(tk.Menu):

    def __init__(self, master=None):
        tk.Menu.__init__(self, master)
        self.master.configure(menu=self)
        self.toplevel_include = None
        self.toplevel_mode = None
        self.selected = []
        self.modes = ["Natal", "Natal"]
        self.settings = tk.Menu(master=self, tearoff=False)
        self.help = tk.Menu(master=self, tearoff=False)
        self.add_cascade(label="Settings", menu=self.settings)
        self.add_cascade(label="Help", menu=self.help)
        self.settings.add_command(
            label="Include",
            command=lambda: self.open_toplevel(
                toplevel=self.toplevel_include,
                func=self.include
            )
        )
        self.settings.add_command(
            label="Mode",
            command=lambda: self.open_toplevel(
                toplevel=self.toplevel_mode,
                func=self.mode
            )
        )
        self.settings.add_command(
            label="Orb Factor",
            command=self.choose_orb_factor
        )
        self.help.add_command(
            label="About",
            command=self.about
        )
        self.help.add_command(
            label="Check for Updates",
            command=self._update
        )
        self.frame = tk.Frame(
            master=self.master,
            width=400,
            height=20
        )
        self.frame.pack()          
        self.button = tk.Button(
            master=self.master,
            text="Start",
            command=self.start
        )
        self.button.pack()

    def start(self):
        threading.Thread(
            target=lambda: Spreadsheet().create_tables(
                file=filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                ),
                selected=self.selected,
                modes=self.modes,
                master=self.frame
            )
        ).start()

    def select_tables(self, checkbuttons: dict = {}):
        self.selected = []
        for i, j in enumerate(ASPECTS):
            if j != "null" and checkbuttons[j][1].get() == "1":
                self.selected.append(j)
        for i in OBJECTS:
            if checkbuttons[i][1].get() == "1":
                self.selected.append(i)
        if checkbuttons["All Aspects"][1].get() == "1":
            self.selected.append("all")
        self.toplevel_include.destroy()
        self.toplevel_include = None

    @staticmethod
    def check_all_buttons(check_all=None, cvar_list: list = [], 
            checkbutton_list: list = []
    ):
        if check_all.get() is True:
            for var, c_button in zip(cvar_list, checkbutton_list):
                var.set(True)
                c_button.configure(variable=var)
        else:
            for var, c_button in zip(cvar_list, checkbutton_list):
                var.set(False)
                c_button.configure(variable=var)

    @staticmethod
    def checkbutton(
            master=None,
            checkbuttons: dict = {},
            text: str = "",
            row: int = 0,
            column: int = 0
    ):
        var = tk.StringVar()
        var.set(value="0")
        cb = tk.Checkbutton(
            master=master,
            text=text,
            variable=var
        )
        cb.grid(row=row, column=column, sticky="w")
        checkbuttons[text] = [cb, var]

    def include(self):
        self.toplevel_include = tk.Toplevel()
        self.toplevel_include.title("Include")
        self.toplevel_include.resizable(width=False, height=False)
        main_frame = tk.Frame(master=self.toplevel_include)
        main_frame.pack()
        left_frame = tk.Frame(
            master=main_frame, 
            bd=1, 
            relief="sunken"
        )
        left_frame.pack(side="left")
        mid_frame = tk.Frame(
            master=main_frame, 
            bd=1, 
            relief="sunken"
        )
        mid_frame.pack(side="left")
        right_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        right_frame.pack(side="left")
        left_frame_label = tk.Label(
            master=left_frame,
            text="Select Aspect Types",
            fg="red"
        )
        left_frame_label.pack()
        mid_frame_label = tk.Label(
            master=mid_frame,
            text="Sum All Aspect Types",
            fg="red"
        )
        mid_frame_label.pack()
        right_frame_label = tk.Label(
            master=right_frame,
            text="Select Planets",
            fg="red"
        )
        right_frame_label.pack()
        left_cb_frame = tk.Frame(master=left_frame)
        left_cb_frame.pack()
        mid_cb_frame = tk.Frame(master=mid_frame)
        mid_cb_frame.pack()
        right_cb_frame = tk.Frame(master=right_frame)
        right_cb_frame.pack()
        checkbuttons = {}
        check_all_1 = tk.BooleanVar()
        check_all_2 = tk.BooleanVar()
        check_uncheck_1 = tk.Checkbutton(
            master=left_cb_frame,
            text="Check/Uncheck All",
            variable=check_all_1)
        check_uncheck_2 = tk.Checkbutton(
            master=right_cb_frame,
            text="Check/Uncheck All",
            variable=check_all_2)
        check_all_1.set(False)
        check_all_2.set(False)
        check_uncheck_1.grid(row=0, column=0, sticky="w")
        check_uncheck_2.grid(row=0, column=0, sticky="w")
        for i, j in enumerate(ASPECTS, 1):
            if j != "null":
                self.checkbutton(
                    master=left_cb_frame,
                    text=j,
                    row=i,
                    column=0,
                    checkbuttons=checkbuttons
                )
        self.checkbutton(
            master=mid_cb_frame,
            text="All Aspects",
            row=0,
            column=0,
            checkbuttons=checkbuttons
        )
        for i, j in enumerate(OBJECTS, 1):
            self.checkbutton(
                master=right_cb_frame,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons
            )
        cvar_list_1 = [i[1] for i in checkbuttons.values()][:11]
        cb_list_1 = [i[0] for i in checkbuttons.values()][:11]
        cvar_list_2 = [i[1] for i in checkbuttons.values()][12:]
        cb_list_2 = [i[0] for i in checkbuttons.values()][12:]
        check_uncheck_1.configure(
            command=lambda: self.check_all_buttons(
                check_all_1,
                cvar_list_1,
                cb_list_1
            )
        )
        check_uncheck_2.configure(
            command=lambda: self.check_all_buttons(
                check_all_2,
                cvar_list_2,
                cb_list_2
            )
        )
        fill_left = tk.Frame(left_cb_frame, height=92)
        fill_left.grid(row=12, column=0)
        fill_mid = tk.Frame(master=mid_cb_frame, height=345)
        fill_mid.grid(row=1, column=0)
        apply_button = tk.Button(
            master=self.toplevel_include,
            text="Apply",
            command=lambda: self.select_tables(checkbuttons)
        )
        apply_button.pack(side="bottom")

    def change_mode(self, variables: list = []):
        self.modes = [i.get() for i in variables]
        self.toplevel_mode.destroy()
        self.toplevel_mode = None

    def mode(self):
        self.toplevel_mode = tk.Toplevel()
        self.toplevel_mode.title("Mode")
        self.toplevel_mode.resizable(width=False, height=False)
        self.modes = []
        main_frame = tk.Frame(master=self.toplevel_mode)
        main_frame.pack()
        variables = []
        for i in range(2):
            frame = tk.Frame(master=main_frame, relief="sunken", bd=1)
            frame.grid(row=0, column=i)
            label = tk.Label(
                master=frame,
                text=f"{i + 1}. Person",
                fg="red"
            )
            label.grid(row=0, column=0)
            var = tk.StringVar()
            var.set("Natal")
            optionmenu = tk.OptionMenu(
                frame, var, "Natal", "Antiscia", "Contra-Antiscia",
            )
            optionmenu["width"] = 15
            optionmenu.grid(row=1, column=0)
            variables.append(var)
        apply = tk.Button(
            master=self.toplevel_mode,
            text="Apply",
            command=lambda: self.change_mode(variables=variables)
        )
        apply.pack()

    @staticmethod
    def open_toplevel(toplevel=None, func=None):
        if toplevel:
            toplevel.destroy()
            func()
        else:
            func()

    def choose_orb_factor(self):
        toplevel = tk.Toplevel()
        toplevel.title("Orb Factor")
        toplevel.resizable(width=False, height=False)
        default_orbs = [
            CONJUNCTION, 
            SEMI_SEXTILE, 
            SEMI_SQUARE,
            SEXTILE, 
            QUINTILE, 
            SQUARE, 
            TRINE,
            SESQUIQUADRATE, 
            BIQUINTILE, 
            QUINCUNX, 
            OPPOSITE
        ]
        orb_entries = []
        for i, j in enumerate(ASPECTS[:-1]):
            aspect_label = tk.Label(master=toplevel, text=f"{j}")
            aspect_label.grid(row=i, column=0, sticky="w")
            equal_to = tk.Label(master=toplevel, text="=")
            equal_to.grid(row=i, column=1, sticky="e")
            orb_entry = tk.Entry(master=toplevel, width=5)
            orb_entry.grid(row=i, column=2)
            orb_entry.insert(0, default_orbs[i])
            orb_entries.append(orb_entry)
        apply_button = tk.Button(
            master=toplevel,
            text="Apply",
            command=lambda: self.change_orb_factors(
                parent=toplevel,
                orb_entries=orb_entries
            )
        )
        apply_button.grid(row=11, column=0, columnspan=3)

    @staticmethod
    def change_orb_factors(parent=None, orb_entries: list = []):
        global CONJUNCTION, SEMI_SEXTILE, SEMI_SQUARE, SEXTILE
        global QUINTILE, SQUARE, TRINE, SESQUIQUADRATE, BIQUINTILE
        global QUINCUNX, OPPOSITE
        CONJUNCTION = int(orb_entries[0].get())
        SEMI_SEXTILE = int(orb_entries[1].get())
        SEMI_SQUARE = int(orb_entries[2].get())
        SEXTILE = int(orb_entries[3].get())
        QUINTILE = int(orb_entries[4].get())
        SQUARE = int(orb_entries[5].get())
        TRINE = int(orb_entries[6].get())
        SESQUIQUADRATE = int(orb_entries[7].get())
        BIQUINTILE = int(orb_entries[8].get())
        QUINCUNX = int(orb_entries[9].get())
        OPPOSITE = int(orb_entries[10].get())
        parent.destroy()

    @staticmethod
    def _update():
        url_1 = "https://raw.githubusercontent.com/dildeolupbiten/" \
                "SynastryParser/master/SynastryParser.py"
        url_2 = "https://raw.githubusercontent.com/dildeolupbiten/" \
                "SynastryParser/master/README.md"
        data_1 = urllib.urlopen(
            url=url_1,
            context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
        data_2 = urllib.urlopen(
            url=url_2,
            context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
        with open(
                file="SynastryParser.py",
                mode="r",
                encoding="utf-8"
        ) as f:
            var_1 = [i.decode("utf-8") for i in data_1]
            var_2 = [i.decode("utf-8") for i in data_2]
            var_3 = [i for i in f]
            if var_1 == var_3:
                msgbox.showinfo(
                    title="Update",
                    message="Program is up-to-date."
                )
            else:
                with open(
                    file="README.md", 
                    mode="w", 
                    encoding="utf-8"
                ) as g:
                    for i in var_2:
                        g.write(i)
                        g.flush()
                with open(
                        file="SynastryParser.py",
                        mode="w",
                        encoding="utf-8"
                ) as h:
                    for i in var_1:
                        h.write(i)
                        h.flush()
                    msgbox.showinfo(
                        title="Update",
                        message="Program is updated."
                    )
                    if os.name == "posix":
                        subprocess.Popen(
                            ["python3", "SynastryParser.py"]
                        )
                        import signal
                        os.kill(os.getpid(), signal.SIGKILL)
                    elif os.name == "nt":
                        subprocess.Popen(
                            ["python", "SynastryParser.py"]
                        )
                        os.system(f"TASKKILL /F /PID {os.getpid()}")
         
    @staticmethod                    
    def callback(event, url):
        webbrowser.open_new(url)
    
    def about(self):
        tl = tk.Toplevel()
        tl.title("About SynastryParser")
        name = "SynastryParser"
        version, _version = "Version:", __version__
        build_date, _build_date = "Built Date:", "02.01.2020"
        update_date, _update_date = "Update Date:", \
            dt.strftime(
                dt.fromtimestamp(os.stat(sys.argv[0]).st_mtime), 
                "%d.%m.%Y"
            )
        developed_by, _developed_by = "Developed By:", \
                                      "Tanberk Celalettin Kutlu"
        contact, _contact = "Contact:", "tckutlu@gmail.com"
        github, _github = "GitHub:", \
            "https://github.com/dildeolupbitenSynastryParser"
        tframe1 = tk.Frame(master=tl, bd="2", relief="groove")
        tframe1.pack(fill="both")
        tframe2 = tk.Frame(master=tl)
        tframe2.pack(fill="both")
        tlabel_title = tk.Label(
            master=tframe1, text=name, font="Arial 25"
        )
        tlabel_title.pack()
        for i, j in enumerate((
            version, build_date, update_date,
            developed_by, contact, github
        )):
            tlabel_info_1 = tk.Label(master=tframe2, text=j,
                                     font="Arial 12", fg="red")
            tlabel_info_1.grid(row=i, column=0, sticky="w")
        for i, j in enumerate((
            _version, _build_date, _update_date,
            _developed_by, _contact, _github
        )):
            if j == _github:
                tlabel_info_2 = tk.Label(master=tframe2, text=j,
                                         font="Arial 12", fg="blue",
                                         cursor="hand2")
                url1 = "https://github.com/dildeolupbiten/SynastryParser"
                tlabel_info_2.bind(
                    "<Button-1>",
                    lambda event: self.callback(event, url1))
            elif j == _contact:
                tlabel_info_2 = tk.Label(master=tframe2, text=j,
                                         font="Arial 12", fg="blue",
                                         cursor="hand2")
                url2 = "mailto://tckutlu@gmail.com"
                tlabel_info_2.bind(
                    "<Button-1>",
                    lambda event: self.callback(event, url2))
            else:
                tlabel_info_2 = tk.Label(master=tframe2, text=j,
                                         font="Arial 12")
            tlabel_info_2.grid(row=i, column=1, sticky="w")


def main():
    root = tk.Tk()
    root.title("SynastryParser")
    root.resizable(height=False, width=False)
    app = App(master=root)
    threading.Thread(target=app.mainloop).run()


if __name__ == "__main__":
    main()
