#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__version__ = "1.2.7"
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
from datetime import datetime as dt

logging.basicConfig(
    format="- %(levelname)s - %(asctime)s - %(message)s",
    level=logging.INFO,
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
    import matplotlib.pyplot as plt
except ModuleNotFoundError:
    os.system("pip3 install matplotlib")
    import matplotlib.pyplot as plt
try:
    import xlwt
except ModuleNotFoundError:
    os.system("pip3 install xlwt")
    import xlwt
try:
    import xlrd
except ModuleNotFoundError:
    os.system("pip3 install xlrd")
    import xlrd
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
try:
    from dateutil import tz
except ModuleNotFoundError:
    os.system("pip3 install python-dateutil")
    from dateutil import tz
try:
    from tzwhere import tzwhere
except ModuleNotFoundError:
    os.system("pip3 install tzwhere")
    from tzwhere import tzwhere

swe.set_ephe_path(os.path.join(os.getcwd(), "Eph"))

URL = "http://cura.free.fr/gauq/Gau_Partners_A_to_M_41832.dat"
FILENAME = URL.split("/")[-1].split(".")[0]
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
    "Neptune",
    "Pluto",
    "True",
    "Chiron",
    "Juno",
    "Asc",
    "MC"
]
ASPECTS = [
    "Conjunction",
    "Semi_Sextile",
    "Semi_Square",
    "Sextile",
    "Quintile",
    "Square",
    "Trine",
    "Sesquiquadrate",
    "Biquintile",
    "Quincunx",
    "Opposite",
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
HOUSES = [f"H{i + 1}" for i in range(12)]
TBL_SS = {
    i: {j: 0 for j in SIGNS}
    for i in SIGNS
}
TBL_HH_PRSNL = {
    i: {j: 0 for j in HOUSES}
    for i in HOUSES
}
TBL_HH_SYNSTRY = {
    i: {j: 0 for j in HOUSES}
    for i in HOUSES
}
TBL_PSPS_PRSNL = {
    sign: {
        house: {
            _sign: {
                _house: 0 for _house in HOUSES
            }
            for _sign in SIGNS
        }
        for house in HOUSES
    }
    for sign in SIGNS
}
TBL_PSPS_SYNSTRY = {
    sign: {
        house: {
            _sign: {
                _house: 0 for _house in HOUSES
            }
            for _sign in SIGNS
        }
        for house in HOUSES
    }
    for sign in SIGNS
}
HSYS = "P"
HOUSE_SYSTEMS = {
    "P": "Placidus",
    "K": "Koch",
    "O": "Porphyrius",
    "R": "Regiomontanus",
    "C": "Campanus",
    "E": "Equal",
    "W": "Whole Signs"
}


def info(s, c, n):
    sys.stdout.write(
        "\r|{}{}| {} %, {} c, {} s estimated, {} c/s, {} s left"
        .format(
            "\u25a0" * int(20 * c / s),
            " " * (20 - int(20 * c / s)),
            int(100 * c / s),
            c,
            int(time.time() - n),
            round(c / (time.time() - n), 1),
            int(s / (c / (time.time() - n))) - int(time.time() - n)
        )
    )
    sys.stdout.flush()


def julday(
        year: int = 0,
        month: int = 0,
        day: int = 0,
        hour: int = 0,
        minute: int = 0,
        second: int = 0
):
    jd = swe.julday(
        year,
        month,
        day,
        hour + (minute / 60) + (second / 3600),
        swe.GREG_CAL
    )
    deltat = swe.deltat(jd)
    return {
        "JD": round(jd + deltat, 6),
        "TT": round(deltat * 86400, 1)
    }


def from_local_to_utc(
        year: int = 0,
        month: int = 0,
        day: int = 0,
        hour: int = 0,
        minute: int = 0,
        lat: float = 0,
        lon: float = 0
):
    tzw = tzwhere.tzwhere()
    timezone = tzw.tzNameAt(lat, lon)
    local_zone = tz.gettz(timezone)
    utc_zone = tz.gettz("UTC")
    if hour == 24:
        hour = 0
    global_time = dt.strptime(
        f"{year}-{month}-{day} {hour}:{minute}:00", "%Y-%m-%d %H:%M:%S"
    )
    local_time = global_time.replace(tzinfo=local_zone)
    utc_time = local_time.astimezone(utc_zone)
    return utc_time.day, utc_time.hour, utc_time.minute, utc_time.second


def convert_raw_data():
    url = URL
    data = [
        [float(j) for j in i.decode().split(",")[1:6]] +
        [float(i.decode().split(",")[-4])] +
        [float(i.decode().split(",")[-5]) * -1]
        for i in urllib.request.urlopen(url)
    ]
    try:
        output = open(f"{FILENAME}.csv", "r+")
    except FileNotFoundError:
        output = open(f"{FILENAME}.csv", "w+")
    readlines = output.readlines()
    count = len(readlines)
    size = len(data)
    now = time.time()
    for i in range(len(readlines), len(data)):
        utc_day, utc_hour, utc_minute, utc_second = from_local_to_utc(
            year=int(data[i][0]),
            month=int(data[i][1]),
            day=int(data[i][2]),
            hour=int(data[i][3]),
            minute=int(data[i][4]),
            lat=data[i][5],
            lon=data[i][6]
        )
        jd = julday(
            year=int(data[i][0]),
            month=int(data[i][1]),
            day=utc_day,
            hour=utc_hour,
            minute=utc_minute,
            second=utc_second
        )
        output.write(f"{jd['JD']},{data[i][-2]},{data[i][-1]}\n")
        output.flush()
        count += 1
        info(s=size, c=count, n=now)
    print()


def rename_csv_file(name):
    file = f"{name}.csv"
    length = len([i for i in open(file, "r")])
    os.rename(file, file.replace("41832", f"{length}"))


def create_control_group():
    url = URL
    data = [
        [int(j) for j in i.decode().split(",")[1:6]][0]
        for i in urllib.request.urlopen(url)
    ]
    males = [data[i] for i in range(0, len(data), 2)]
    females = [data[i] for i in range(1, len(data), 2)]
    temp_females = [i for i in females]
    count = 2
    now = time.time()
    check_list = []
    try:
        with open(f"{FILENAME}.csv", "r") as case:
            logging.info("Creating control group...")
            case_lines = case.readlines()
            size = len(case_lines)
            with open(f"Random_{FILENAME}.csv", "w") \
                    as control:
                for i, male in enumerate(males):
                    for j, female in enumerate(temp_females):
                        if females[i] == female and j != 0:
                            try:
                                check_list.append(female)
                                index = [
                                    k * 2 + 1 for k in range(len(females))
                                    if females[k] == female
                                ][check_list.count(female)]
                                control.write(f"{case_lines[i * 2]}")
                                control.write(f"{case_lines[index]}")
                                control.flush()
                                temp_females.pop(j)
                                break
                            except IndexError:
                                pass
                    info(s=size, c=count, n=now)
                    count += 2
            rename_csv_file(name=f"Random_{FILENAME}")
            logging.info("Completed creating control group.")
    except FileNotFoundError:
        logging.info(f"{FILENAME}.csv is not found.")


def split_gauquelin_data():
    url = URL
    data = [
        [int(j) for j in i.decode().split(",")[1:6]][0]
        for i in urllib.request.urlopen(url)
    ]
    males = [data[i] for i in range(0, len(data), 2)]
    try:
        with open(f"{FILENAME}.csv", "r") as case:
            with open(f"{FILENAME}_Before_1900.csv", "w") as before:
                with open(f"{FILENAME}_After_1900.csv", "w") as after:
                    logging.info("Splitting Gauquelin Data...")
                    case_lines = case.readlines()
                    size = len(case_lines)
                    count = 0
                    now = time.time()
                    for i in range(len(males)):
                        if males[i] < 1900:
                            before.write(f"{case_lines[i * 2]}")
                            before.write(f"{case_lines[i * 2 + 1]}")
                            before.flush()
                        elif males[i] >= 1900:
                            after.write(f"{case_lines[i * 2]}")
                            after.write(f"{case_lines[i * 2 + 1]}")
                            after.flush()
                        count += 2
                        info(s=size, c=count, n=now)
            rename_csv_file(name=f"{FILENAME}_Before_1900")
            rename_csv_file(name=f"{FILENAME}_After_1900")
            logging.info("Completed splitting control group.")
    except FileNotFoundError:
        logging.info(f"{FILENAME}.csv is not found.")


class Chart:
    PLANET_DICT = {
        OBJECTS[0]: swe.SUN,
        OBJECTS[1]: swe.MOON,
        OBJECTS[2]: swe.MERCURY,
        OBJECTS[3]: swe.VENUS,
        OBJECTS[4]: swe.MARS,
        OBJECTS[5]: swe.JUPITER,
        OBJECTS[6]: swe.SATURN,
        OBJECTS[7]: swe.URANUS,
        OBJECTS[8]: swe.NEPTUNE,
        OBJECTS[9]: swe.PLUTO,
        OBJECTS[10]: swe.TRUE_NODE,
        OBJECTS[11]: swe.CHIRON,
        OBJECTS[12]: swe.JUNO
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
        planet_positions = []
        house_positions = []
        for i in range(12):
            house = [
                int(self.house_pos()[0][i][0]),
                self.house_pos()[0][i][-1],
                float(self.house_pos()[0][i][1]),
            ]
            house_positions.append(house)
        hp = [j[-1] for j in house_positions]
        for key, value in self.PLANET_DICT.items():
            planet = self.planet_pos(planet=value)
            house = 0
            for i in range(12):
                if i != 11:
                    if hp[i] < planet[1] < hp[i + 1]:
                        house = i + 1
                        break                        
                    elif hp[i] < planet[1] > hp[i + 1] \
                            and hp[i] - hp[i + 1] > 240: 
                        house = i + 1
                        break
                    elif hp[i] > planet[1] < hp[i + 1] \
                            and hp[i] - hp[i + 1] > 240: 
                        house = i + 1
                        break
                else:
                    if hp[i] < planet[1] < hp[0]:
                        house = i + 1
                        break                        
                    elif hp[i] < planet[1] > hp[0] \
                            and hp[i] - hp[0] > 240: 
                        house = i + 1
                        break
                    elif hp[i] > planet[1] < hp[0] \
                            and hp[i] - hp[0] > 240: 
                        house = i + 1
                        break
            planet_info = [
                key,
                planet[0],
                planet[1],
                f"H{house}"
            ]
            planet_positions.append(planet_info)
        asc = house_positions[0] + ["H1"]
        asc[0] = "Asc"
        mc = house_positions[9] + ["H10"]
        mc[0] = "MC"
        planet_positions.extend([asc, mc])
        return planet_positions, house_positions


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


def create_table(
        _x: list = [], _y: list = [], obj: str = "",
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


def change_mode(_info: list = [], mode: dict = {}):
    for i, j in enumerate(_info):
        _info[i] = [_info[i][0]] + \
                   [mode[_info[i][1]]] + \
                   [
                       SST[mode[_info[i][1]]] +
                       30 - _info[i][2] % 30
                   ]
    return _info


def change_values(selection: str = "", _info: list = []):
    if selection == "Antiscia":
        _info = change_mode(_info=_info, mode=ANTISCIA)
    elif selection == "Contra-Antiscia":
        _info = change_mode(_info=_info, mode=CONTRA_ANTISCIA)
    return _info


def activate_selections(
        selected: list = [], modes: list = [],
        male: list = [], female: list = []
):
    global x, y
    x, y = male, female
    x = change_values(selection=modes[0], _info=x)
    y = change_values(selection=modes[1], _info=y)
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
                    code = f"create_table(_x=x, _y=y, obj='{i}', " \
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


def count_aspects_1(arg=(), file_list=[], count=1):
    result = ()
    for i in file_list[count:]:
        for j, k in zip(arg, i):
            aspect = {}
            for (m, n), (o, p) in zip(j.items(), k.items()):
                aspect[m] = sum_arrays(n, p)
            result += aspect,
        if count < len(file_list) - 1:
            count += 1
            return count_aspects_1(result, file_list, count)
        else:
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


def r_aspect_dist_acc_to_objects(
        file: list = [],
        selected: list = [],
        modes: list = []
):
    file_list = []
    split_list = []
    n = len(file)
    for i in range(0, n, 2):
        male = [float(j) for j in file[i].strip().split(",")]
        female = [float(j) for j in file[i + 1].strip().split(",")]
        splitted = activate_selections(
            selected=selected,
            male=[j[:-1] for j in Chart(*male, HSYS).patterns()[0]],
            female=[j[:-1] for j in Chart(*female, HSYS).patterns()[0]],
            modes=modes
        )
        split_list.append(splitted[1])
        file_list.append(splitted[0])
    item1 = file_list[0]
    return count_aspects_1(item1, file_list, 1), \
        count_aspects_2(split_list, selected)


def synastry_pos(chart1: list = [], chart2: list = []):
    hp = [j[-1] for j in chart2[1]]
    for ind, planet in enumerate(chart1[0]):
        house = 0
        for i in range(12):
            if i != 11:
                if hp[i] < planet[1] < hp[i + 1]:
                    house = i + 1
                    break                        
                elif hp[i] < planet[1] > hp[i + 1] \
                        and hp[i] - hp[i + 1] > 240: 
                    house = i + 1
                    break
                elif hp[i] > planet[1] < hp[i + 1] \
                        and hp[i] - hp[i + 1] > 240: 
                    house = i + 1
                    break
            else:
                if hp[i] < planet[1] < hp[0]:
                    house = i + 1
                    break                        
                elif hp[i] < planet[1] > hp[0] \
                        and hp[i] - hp[0] > 240: 
                    house = i + 1
                    break
                elif hp[i] > planet[1] < hp[0] \
                        and hp[i] - hp[0] > 240: 
                    house = i + 1
                    break
        chart1[0][ind].append(f"H{house}")
    return chart1[0]


def r_syn_object_dist_acc_to_signs_or_houses(
        file: str = "",
        arg1: str = "",
        arg2: str = "",
        table: dict = {},
):
    keys = list(Chart.PLANET_DICT.keys())
    keys.extend(["Asc", "MC"])
    with open(file, "r") as f:
        readlines = f.readlines()
        size = len(readlines)
        now = time.time()
        count = 0
        for i in range(0, size, 2):
            old_male = Chart(
                *[
                    float(col)
                    for col in readlines[i][:-1].split(",")
                ],
                HSYS
            ).patterns()
            old_female = Chart(
                *[
                    float(col)
                    for col in readlines[i + 1][:-1].split(",")
                ],
                HSYS
            ).patterns()
            male = synastry_pos(
                old_male, old_female)[keys.index(arg1)][-1]
            female = synastry_pos(
                old_female, old_male)[keys.index(arg2)][-1]
            table[male][female] += 1
            count += 2
            info(s=size, c=count, n=now)
        print()


def r_object_dist_acc_to_signs_or_houses(
        file: str = "",
        arg1: str = "",
        arg2: str = "",
        table: dict = {},
        index: int = 0
):
    keys = list(Chart.PLANET_DICT.keys())
    keys.extend(["Asc", "MC"])
    with open(file, "r") as f:
        readlines = f.readlines()
        size = len(readlines)
        now = time.time()
        count = 0
        for i in range(0, size, 2):
            male = Chart(
                *[float(col) for col in readlines[i][:-1].split(",")],
                HSYS
            ).patterns()[0][keys.index(arg1)][index]
            female = Chart(
                *[float(col) for col in readlines[i + 1][:-1].split(",")],
                HSYS
            ).patterns()[0][keys.index(arg2)][index]
            table[male][female] += 1
            count += 2
            info(s=size, c=count, n=now)
        print()


def r_object_sign_dist_acc_to_houses(
        file: str = "",
        arg1: str = "",
        arg2: str = "",
):
    keys = list(Chart.PLANET_DICT.keys())
    keys.extend(["Asc", "MC"])
    with open(file, "r") as f:
        readlines = f.readlines()
        size = len(readlines)
        now = time.time()
        count = 0
        for i in range(0, size, 2):
            old_male = Chart(
                *[float(col) for col in readlines[i][:-1].split(",")],
                HSYS
            ).patterns()
            old_female = Chart(
                *[float(col) for col in readlines[i + 1][:-1].split(",")],
                HSYS
            ).patterns()
            male = synastry_pos(
                old_male, old_female)[keys.index(arg1)]
            female = synastry_pos(
                old_female, old_male)[keys.index(arg2)]
            TBL_PSPS_PRSNL[male[1]][male[-1]][female[1]][female[-1]] += 1
            count += 2
            info(s=size, c=count, n=now)
        print()


def r_syn_object_sign_dist_acc_to_houses(
        file: str = "",
        arg1: str = "",
        arg2: str = "",
):
    keys = list(Chart.PLANET_DICT.keys())
    keys.extend(["Asc", "MC"])
    with open(file, "r") as f:
        readlines = f.readlines()
        size = len(readlines)
        now = time.time()
        count = 0
        for i in range(0, size, 2):
            old_male = Chart(
                *[float(col) for col in readlines[i][:-1].split(",")],
                HSYS
            ).patterns()
            old_female = Chart(
                *[float(col) for col in readlines[i + 1][:-1].split(",")],
                HSYS
            ).patterns()
            male = synastry_pos(old_male, old_female)[keys.index(arg1)]
            female = synastry_pos(old_female, old_male)[keys.index(arg2)]
            TBL_PSPS_SYNSTRY[male[1]][male[-1]][female[1]][female[-1]] += 1
            count += 2
            info(s=size, c=count, n=now)
        print()


def font(name: str = "Arial", bold: bool = False):
    f = xlwt.Font()
    f.name = name
    f.bold = bold
    return f


class Spreadsheet(xlwt.Workbook):
    size = 5
    obj = None

    def __init__(
            self,
            arg1: str = "",
            arg2: str = "",
            arg3: str = "",
            arg4: str = ""
    ):
        xlwt.Workbook.__init__(self)
        self.arg1 = arg1
        self.arg2 = arg2
        self.arg3 = arg3
        self.arg4 = arg4
        self.letters = [
            "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"
        ]
        self.sheet = self.add_sheet("Sheet1")
        self.style = xlwt.XFStyle()
        self.alignment = xlwt.Alignment()
        self.alignment.horz = xlwt.Alignment.HORZ_CENTER
        self.alignment.vert = xlwt.Alignment.VERT_CENTER
        self.style.alignment = self.alignment

    def w_aspect_dist_acc_to_objects(
            self, table: dict = {}, aspect: str = ""
    ):
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

    def create_tables(
            self, files: list = [], selected: list = [],
            modes: list = [], num: int = 0, file: str = "",
            number_of_records: int = 0
    ):
        read = r_aspect_dist_acc_to_objects(files, selected, modes)
        self.style.font = font(bold=True)
        self.sheet.write(
            r=0, c=0, label="File", style=self.style
        )
        self.style.font = font(bold=False)
        self.sheet.write_merge(
            r1=0, c1=1, r2=0, c2=4,
            label=f"{os.path.split(file)[-1]}", style=self.style
        )
        self.style.font = font(bold=True)
        self.sheet.write_merge(
            r1=1, c1=0, r2=1, c2=1, label="Number of Records",
            style=self.style
        )
        self.style.font = font(bold=False)
        self.sheet.write(
            r=1, c=2,
            label=f"{number_of_records}", style=self.style
        )
        self.style.font = font(bold=True)
        self.sheet.write_merge(
            r1=2, c1=0, r2=3, c2=0, label="Mode", style=self.style
        )
        self.sheet.write(r=2, c=1, label="Males", style=self.style)
        self.sheet.write(r=2, c=2, label="Females", style=self.style)
        self.style.font = font(bold=False)
        self.sheet.write(r=3, c=1, label=modes[0], style=self.style)
        self.sheet.write(r=3, c=2, label=modes[1], style=self.style)
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
                    self.w_aspect_dist_acc_to_objects(
                        tables[-1],
                        "all",
                    )
                count = 0
                for index, aspect in enumerate(ASPECTS):
                    if aspect in selected:
                        self.w_aspect_dist_acc_to_objects(
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
                                        label="Females'\n" + j +
                                              "\nin Signs",
                                        style=self.style
                                    )
                                    self.sheet.write_merge(
                                        r1=self.size,
                                        r2=self.size,
                                        c1=2,
                                        c2=13,
                                        label="Males' " + i +
                                              " in Signs",
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
                self.save(f"part{str(num).zfill(3)}.xlsx")
            else:
                pass
        except TypeError:
            pass

    def w_object_dist_acc_to_signs_or_houses(
            self,
            table: dict = {},
            modes: list = [],
            selected_obj: list = [],
            selection: str = "",
            file: str = "",
            number_of_records: int = 0
    ):
        self.style.font = font(bold=True)
        self.sheet.write(
            r=0, c=0, label="File", style=self.style
        )
        self.style.font = font(bold=False)
        self.sheet.write_merge(
            r1=0, c1=1, r2=0, c2=4,
            label=f"{os.path.split(file)[-1]}", style=self.style
        )
        self.style.font = font(bold=True)
        self.sheet.write_merge(
            r1=1, c1=0, r2=1, c2=1, label="Number of Records",
            style=self.style
        )
        self.style.font = font(bold=False)
        self.sheet.write(
            r=1, c=2,
            label=f"{number_of_records}", style=self.style
        )
        if selection == "Sign":
            arg = SIGNS
            _row = 2
        else:
            arg = HOUSES
            _row = 3
            self.style.font = font(bold=True)
            self.sheet.write_merge(
                r1=2, c1=0, r2=2, c2=1, label="House System",
                style=self.style
            )
            self.style.font = font(bold=False)
            self.sheet.write(
                r=2, c=2, label=HOUSE_SYSTEMS[HSYS], style=self.style
            )
            self.style.font = font(bold=True)
        self.style.font = font(bold=True)
        self.sheet.write_merge(
            r1=_row, c1=0, r2=_row + 1, c2=0, label="Mode", style=self.style
        )
        self.sheet.write(r=_row, c=1, label="Males", style=self.style)
        self.sheet.write(r=_row, c=2, label="Females", style=self.style)
        self.style.font = font(bold=False)
        self.sheet.write(r=_row + 1, c=1, label=modes[0], style=self.style)
        self.sheet.write(r=_row + 1, c=2, label=modes[1], style=self.style)
        self.style.font = font(bold=True)
        if selection == "Personal_House":
            male_label = "Males' " + selected_obj[0] + " in Houses"
            female_label = "Females'\n" + selected_obj[1] + "\nin Houses"
        elif selection == "Synastry_House":
            male_label = "Males' " + selected_obj[0] + \
                         " in Females' Houses"
            female_label = "Females'\n" + selected_obj[1] + \
                           "\nin Males'\nHouses"
        else:
            male_label = "Males' " + selected_obj[0] + " in Signs"
            female_label = "Females'\n" + selected_obj[1] + "\nin Signs"
        self.sheet.write_merge(
            r1=_row + 3, c1=2, r2=_row + 3, c2=13,
            label=male_label,
            style=self.style
        )
        self.sheet.write_merge(
            r1=_row + 5, c1=0, r2=_row + 16, c2=0,
            label=female_label,
            style=self.style
        )
        for i, i_ in enumerate(arg):
            self.sheet.write(r=_row + 4, c=i + 2, label=i_, style=self.style)
            self.sheet.write(r=_row + i + 5, c=1, label=i_, style=self.style)
        self.sheet.write(r=_row + 4, c=14, label="Total", style=self.style)
        self.sheet.write(r=_row + 17, c=1, label="Total", style=self.style)
        self.style.font = font(bold=False)
        row = _row + 5
        column = 2
        for keys, values in table.items():
            r = 0
            for subkeys, subvalues in values.items():
                self.sheet.write(
                    r=row + r,
                    c=column,
                    label=subvalues,
                    style=self.style
                )
                r += 1
            column += 1
        column = 2
        row = _row + 5
        for i in self.letters:
            self.sheet.write(
                _row + 17, column,
                xlwt.Formula(
                    f"SUM({i}{_row + 6}:{i}{_row + 17})"
                ),
                style=self.style
            )

            self.sheet.write(
                row, 14,
                xlwt.Formula(
                    f"SUM(C{row + 1}:N{row + 1})"
                ),
                style=self.style
            )
            column += 1
            row += 1
        self.sheet.write(
            _row + 17, 14,
            xlwt.Formula(
                f"SUM(C{_row + 18}:N{_row + 18})"
            ),
            style=self.style
        )
        self.save(
            f"{selection}_"
            f"Male_{self.arg1}_Female_{self.arg2}.xlsx"
        )

    def w_object_sign_dist_acc_to_houses(
            self,
            table: dict = {},
            modes: list = [],
            selected_obj: list = [],
            selected_sign: list = [],
            selection: str = "",
            file: str = "",
            number_of_records: int = 0
    ):
        self.style.font = font(bold=True)
        self.sheet.write(
            r=0, c=0, label="File", style=self.style
        )
        self.style.font = font(bold=False)
        self.sheet.write_merge(
            r1=0, c1=1, r2=0, c2=4,
            label=f"{os.path.split(file)[-1]}", style=self.style
        )
        self.style.font = font(bold=True)
        self.sheet.write_merge(
            r1=1, c1=0, r2=1, c2=1, label="Number of Records",
            style=self.style
        )
        self.style.font = font(bold=False)
        self.sheet.write(
            r=1, c=2,
            label=f"{number_of_records}", style=self.style
        )
        self.style.font = font(bold=True)
        self.sheet.write_merge(
            r1=2, c1=0, r2=2, c2=1, label="House System", style=self.style)
        self.style.font = font(bold=False)
        self.sheet.write(
            r=2, c=2, label=HOUSE_SYSTEMS[HSYS], style=self.style
        )
        self.style.font = font(bold=True)
        self.sheet.write_merge(
            r1=3, c1=0, r2=4, c2=0, label="Mode", style=self.style
        )
        self.sheet.write(r=3, c=1, label="Males", style=self.style)
        self.sheet.write(r=3, c=2, label="Females", style=self.style)
        self.style.font = font(bold=False)
        self.sheet.write(r=4, c=1, label=modes[0], style=self.style)
        self.sheet.write(r=4, c=2, label=modes[1], style=self.style)
        self.style.font = font(bold=True)
        if selection == "Synastry_Object_Sign":
            male_label = f"Males' {selected_obj[0]} " \
                         f"({selected_sign[0]}) in Females' Houses"
            female_label = f"Females'\n{selected_obj[1]}\n" \
                           f"({selected_sign[1]})\nin Males'\nHouses"
            filename = "Synastry"
        else:
            male_label = f"Males' {selected_obj[0]} " \
                         f"({selected_sign[0]}) in Houses"
            female_label = f"Females'\n{selected_obj[1]}\n" \
                           f"({selected_sign[1]})\nin Houses"
            filename = "Personal"
        self.sheet.write_merge(
            r1=6,
            c1=2,
            r2=6,
            c2=13,
            label=male_label,
            style=self.style
        )
        self.sheet.write_merge(
            r1=8,
            c1=0,
            r2=19,
            c2=0,
            label=female_label,
            style=self.style
        )
        for i, i_ in enumerate(HOUSES):
            self.sheet.write(r=7, c=i + 2, label=i_, style=self.style)
            self.sheet.write(r=i + 8, c=1, label=i_, style=self.style)
        self.sheet.write(r=7, c=14, label="Total", style=self.style)
        self.sheet.write(r=20, c=1, label="Total", style=self.style)
        self.style.font = font(bold=False)
        column = 2
        for i in HOUSES:
            row = 8
            for j in HOUSES:
                value = table[selected_sign[0]][i][selected_sign[-1]][j]
                self.sheet.write(
                    r=row,
                    c=column,
                    label=value,
                    style=self.style
                )
                row += 1
            column += 1
        for i in range(12):
            self.sheet.write(
                20, i + 2,
                xlwt.Formula(
                    f"SUM({self.letters[i]}8:{self.letters[i]}19)"
                ),
                style=self.style)
            self.sheet.write(
                i + 8, 14,
                xlwt.Formula(
                    f"SUM(C{i + 9}:N{i + 9})"
                ),
                style=self.style)
        self.sheet.write(
            20, 14,
            xlwt.Formula(
                f"SUM(O9:O20)"
            ),
            style=self.style)
        self.save(
            f"{filename}_Male_{self.arg1}_{self.arg3}_"
            f"Female_{self.arg2}_{self.arg4}.xlsx"
        )


class Plot:

    @staticmethod
    def __plt(_x_label="", _y_label="", _title="x / y"):
        plt.grid(
            True,
            color="black",
            linestyle="-",
            linewidth=1
        )
        plt.xlabel(_x_label)
        plt.ylabel(_y_label)
        plt.legend()
        plt.show()

    @classmethod
    def plot(cls, *args):
        n = 0
        for i, j in zip(["Male", "Female"], ["blue", "red"]):
            plt.plot(
                args[n],
                args[n + 1],
                label=i,
                color=j
            )
            n += 2
        cls.__plt(_x_label=args[-1], _y_label="Number of people")

    @classmethod
    def bar(cls, *args):
        n = 0
        for i in range(len(args) // 2):
            plt.bar(
                args[n],
                args[n + 1],
                label="Couple"
            )
            n += 2
        cls.__plt(_x_label="Age Difference", _y_label="Number of couple")


def frequency(l: list = [], d: dict = {}):
    for i in l:
        if i in d.keys():
            d[i] += 1
        else:
            d[i] = 1


def year_long_lati_frequency(
        name: str = "",
        index: int = 0
):
    url = URL
    data = [
        [float(j) for j in i.decode().split(",")[1:6]] +
        [float(i.decode().split(",")[-4])] +
        [float(i.decode().split(",")[-5]) * -1]
        for i in urllib.request.urlopen(url)
    ]
    male_list = sorted(
        [
            data[i][index]
            for i in range(0, len(data), 2)
        ]
    )
    female_list = sorted(
        [
            data[i][index]
            for i in range(1, len(data), 2)
        ]
    )
    male_dict, female_dict = {}, {}
    frequency(l=male_list, d=male_dict)
    frequency(l=female_list, d=female_dict)
    with open(f"{name}Frequency.txt", "w") as f:
        f.write(
            f"|{'Male'.center(29)}|{'Female'.center(29)}|\n"
            f"|{name[:4].center(14)}|{'Count'.center(14)}|"
            f"{name[:4].center(14)}|{'Count'.center(14)}|\n"
        )
        for (i, j), (k, m) in zip(male_dict.items(), female_dict.items()):
            f.write(
                f"|{str(i).center(14)}|{str(j).center(14)}"
                f"|{str(k).center(14)}|{str(m).center(14)}|\n"
            )
            f.flush()
    Plot.plot(
        list(male_dict.keys()),
        list(male_dict.values()),
        list(female_dict.keys()),
        list(female_dict.values()),
        name
    )


def age_differences_frequency():
    url = URL
    age_diffs = []
    age_diff_freq = {i: 0 for i in range(42)}
    data = [
        [str(j) for j in i.decode().split(",")[1:6]] +
        [str(i.decode().split(",")[-4])] +
        [float(i.decode().split(",")[-5]) * -1]
        for i in urllib.request.urlopen(url)
    ]
    for i in range(len(data)):
        if data[i][3] == "24":
            data[i][3] = "0"
    male_list = [
        dt.strptime(
            ".".join(data[i][0:3]) + " " +
            ":".join(data[i][3:5]),
            "%Y.%m.%d %H:%M"
        )
        for i in range(0, len(data), 2)
    ]
    female_list = [
        dt.strptime(
            ".".join(data[i][0:3]) + " " +
            ":".join(data[i][3:5]),
            "%Y.%m.%d %H:%M"
        )
        for i in range(1, len(data), 2)
    ]
    f = open("age_diff.txt", "w")
    f.write(f"|    COUPLES   | AGE DIFFERENCE (YEAR) |\n")
    count = 0
    for i in range(len(male_list)):
        if female_list[i] > male_list[i]:
            diff = round(
                (female_list[i] - male_list[i]).total_seconds() /
                (60 * 60 * 24 * 365.2422), 2
            )
        else:
            diff = round(
                (male_list[i] - female_list[i]).total_seconds() /
                (60 * 60 * 24 * 365.2422), 2
            )
        f.write(
            f"| Couple {str(i + 1).center(5)} | {str(diff).center(21)} |\n"
        )
        count += diff
        age_diffs.append(diff)
        for index, key in enumerate(list(age_diff_freq.keys())):
            if index != len(list(age_diff_freq.keys())) - 1:
                if list(age_diff_freq.keys())[index] <= diff < \
                        list(age_diff_freq.keys())[index + 1]:
                    age_diff_freq[key] += 1
    f.write(f"\n\nMean Difference: {count / len(male_list)}\n")
    f.write(f"Maximum Difference: {max(age_diffs)}\n")
    f.write(f"Minimum Difference: {min(age_diffs)}\n")
    f.flush()
    g = open("age_diff_freq.txt", "w")
    g.write("| AGE DIFFERENCE (YEAR) | TOTAL |\n")
    for keys, values in age_diff_freq.items():
        g.write(f"|{str(keys).center(23)}|{str(values).center(7)}|\n")
        g.flush()
    Plot.bar(
        list(age_diff_freq.keys()),
        list(age_diff_freq.values())
    )


class App(tk.Menu):

    def __init__(self, master=None):
        tk.Menu.__init__(self, master)
        self.master.configure(menu=self)
        self.t_aspect_general = None
        self.t_aspect_dist_acc_to_objects = None
        self.t_object_dist_acc_to_signs = None
        self.t_object_dist_acc_to_houses = None
        self.t_object_sign_dist_acc_to_houses = None
        self.t_mode = None
        self.t_orb = None
        self.t_hsys = None
        self.selected = []
        self.selected_obj = []
        self.selected_sign = []
        self.selection = ""
        self.modes = ["Natal", "Natal"]
        self.records = tk.Menu(master=self, tearoff=False)
        self.convert = tk.Menu(master=self, tearoff=False)
        self.create = tk.Menu(master=self, tearoff=False)
        self.frequency = tk.Menu(master=self, tearoff=False)
        self.settings = tk.Menu(master=self, tearoff=False)
        self.tables = tk.Menu(master=self, tearoff=False)
        self.help = tk.Menu(master=self, tearoff=False)
        self.add_cascade(label="Records", menu=self.records)
        self.add_cascade(label="Frequency", menu=self.frequency)
        self.add_cascade(label="Settings", menu=self.settings)
        self.add_cascade(label="Tables", menu=self.tables)
        self.add_cascade(label="Help", menu=self.help)
        self.records.add_command(
            label=f"Convert Gauquelin Data",
            command=lambda: threading.Thread(
                target=convert_raw_data
            ).start()
        )
        self.records.add_command(
            label="Split Gauquelin Data",
            command=lambda: threading.Thread(
                target=split_gauquelin_data
            ).start()
        )
        self.records.add_command(
            label="Create Control Group",
            command=lambda: threading.Thread(
                target=create_control_group
            ).start()
        )
        self.frequency.add_command(
            label="Longitude Frequency",
            command=lambda: year_long_lati_frequency(
                name="Longitude", index=-1
            )
        )
        self.frequency.add_command(
            label="Latitude Frequency",
            command=lambda: year_long_lati_frequency(
                name="Latitude", index=-2
            )
        )
        self.frequency.add_command(
            label="Year Frequency",
            command=lambda: year_long_lati_frequency(
                name="Year", index=0
            )
        )
        self.frequency.add_command(
            label="Age Difference Frequency",
            command=age_differences_frequency
        )
        self.tables.add_command(
            label="General Aspect distributions",
            command=lambda: self.open_toplevel(
                toplevel=self.t_aspect_general,
                func=self.aspect_general
            )
        )
        self.tables.add_command(
            label="Aspect distributions of objects",
            command=lambda: self.open_toplevel(
                toplevel=self.t_aspect_dist_acc_to_objects,
                func=self.aspect_dist_acc_to_objects
            )
        )
        self.tables.add_command(
            label="Sign positions of objects",
            command=lambda: self.open_toplevel(
                toplevel=self.t_object_dist_acc_to_signs,
                func=self.object_dist_acc_to_signs
            )
        )
        self.tables.add_command(
            label="House positions of objects (Personal)",
            command=lambda: self.open_toplevel(
                toplevel=self.t_object_dist_acc_to_houses,
                func=lambda: self.object_dist_acc_to_houses(
                    title="House positions of objects (Personal)",
                    selection="Personal_House"
                )
            )
        )
        self.tables.add_command(
            label="House positions of objects-signs (Personal)",
            command=lambda: self.open_toplevel(
                toplevel=self.t_object_sign_dist_acc_to_houses,
                func=lambda: self.object_sign_dist_acc_to_houses(
                    title="House positions of objects-signs (Personal)",
                    selection="Object_Sign"
                )
            )
        )
        self.tables.add_command(
            label="House positions of objects (Synastry)",
            command=lambda: self.open_toplevel(
                toplevel=self.t_object_dist_acc_to_houses,
                func=lambda: self.object_dist_acc_to_houses(
                    title="House positions of objects (Synastry)",
                    selection="Synastry_House"
                )
            )
        )
        self.tables.add_command(
            label="House positions of objects-signs (Synastry)",
            command=lambda: self.open_toplevel(
                toplevel=self.t_object_sign_dist_acc_to_houses,
                func=lambda: self.object_sign_dist_acc_to_houses(
                    title="House positions of objects-signs (Synastry)",
                    selection="Synastry_Object_Sign"
                )
            )
        )
        self.settings.add_command(
            label="House System",
            command=self.create_hsys_checkbuttons
        )
        self.settings.add_command(
            label="Mode",
            command=lambda: self.open_toplevel(
                toplevel=self.t_mode,
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
            command=self.update_script
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
            command=lambda: threading.Thread(
                target=self.start,
                args=(
                    self.selected,
                    self.selected_obj,
                    self.selected_sign
                )
            ).start()
        )
        self.style = xlwt.XFStyle()
        self.alignment = xlwt.Alignment()
        self.alignment.horz = xlwt.Alignment.HORZ_CENTER
        self.alignment.vert = xlwt.Alignment.VERT_CENTER
        self.style.alignment = self.alignment
        self.button.pack()

    @staticmethod
    def get_data(sheet):
        data = []
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                data.append(([row, col], sheet.cell_value(row, col)))
        return data

    def special_cells(self, new_sheet, i):
        self.style.font = font(bold=True)
        if "Orb" in i[1]:
            new_sheet.write_merge(
                r1=i[0][0],
                r2=i[0][0],
                c1=i[0][1],
                c2=i[0][1] + 15,
                label=i[1],
                style=self.style
            )
        elif "Mode" in i[1]:
            new_sheet.write_merge(
                r1=i[0][0],
                r2=i[0][0] + 1,
                c1=i[0][1],
                c2=i[0][1],
                label=i[1],
                style=self.style
            )
        elif "ALL" in i[1]:
            self.style.font = font(bold=True)
            new_sheet.write_merge(
                r1=i[0][0],
                r2=i[0][0],
                c1=i[0][1],
                c2=i[0][1] + 15,
                label=i[1],
                style=self.style
            )        
        elif "csv" in i[1]:
            self.style.font = font(bold=False)
            new_sheet.write_merge(
                r1=i[0][0],
                r2=i[0][0],
                c1=i[0][1],
                c2=i[0][1] + 4,
                label=i[1],
                style=self.style
            )
        elif "Number of Records" in i[1]:
            new_sheet.write_merge(
                r1=i[0][0],
                r2=i[0][0],
                c1=i[0][1],
                c2=i[0][1] + 1,
                label=i[1],
                style=self.style
            )
        elif i[0][0] in [i__ for i__ in range(23, 232, 15)] and \
                self.selection != "Aspect-General":
            new_sheet.write_merge(
                r1=i[0][0],
                r2=i[0][0],
                c1=2,
                c2=13,
                label=i[1],
                style=self.style
            )
        elif i[0][0] in [i__ for i__ in range(25, 250, 15)] \
                and i[0][1] == 0 and self.selection != "Aspect-General":
            new_sheet.write_merge(
                r1=i[0][0],
                r2=i[0][0] + 11,
                c1=0,
                c2=0,
                label=i[1],
                style=self.style
            )
        elif i[0][0] == 1:
            self.style.font = font(bold=False)
            new_sheet.write(*i[0], i[1], style=self.style)
            self.style.font = font(bold=True)
        else:
            if "Natal" in i[1]:
                self.style.font = font(bold=False)
            else:
                self.style.font = font(bold=True)
            new_sheet.write(*i[0], i[1], style=self.style)
        self.style.font = font(bold=False)

    def run(self):
        datas = lambda: [
            self.get_data(
                xlrd.open_workbook(file).sheet_by_name("Sheet1")
            )
            for file in sorted(os.listdir("."))
            if file.startswith("part") and file.endswith("xlsx")
        ]
        if self.selection == "Aspect-General":
            msg = "Calculation of 'General aspect distributions' "\
                "is completed."
        else:
            msg = "Calculation of 'Aspect distributions of objects' "\
                "is completed."
        index = 2
        s = len(datas())
        c = 1
        t = time.time()
        while len(datas()) != 1:
            _datas = datas()
            new_file = xlwt.Workbook()
            new_sheet = new_file.add_sheet("Sheet1")
            control_list = []
            for i, j in zip(_datas[0], _datas[1]):
                if i[0] == j[0]:
                    if i[1] != "" and j[1] != "":
                        if isinstance(i[1], float) and \
                                isinstance(j[1], float):
                            new_sheet.write(
                                *i[0],
                                i[1] + j[1],
                                style=self.style
                            )
                        else:
                            self.special_cells(new_sheet, i)
                    elif (i[1] != "" and j[1] == "") or \
                            (i[1] == "" and j[1] != ""):
                        new_sheet.write(*i[0], i[1], style=self.style)
                    elif i[1] == "" and j[1] != "":
                        new_sheet.write(*j[0], j[1], style=self.style)
                    if i[0] not in control_list:
                        control_list.append(i[0])
                    if j[0] not in control_list:
                        control_list.append(j[0])
            os.remove(f"part001.xlsx")
            os.remove(f"part{str(index).zfill(3)}.xlsx")
            new_file.save(f"part001.xlsx")
            index += 1
            c += 1
            info(s=s, c=c, n=t)
        else:
            print()
            logging.info("Completed merging the separated files.")
            if self.selection == "Aspect-Detailed":
                os.rename("part001.xlsx", f"{'_'.join(self.selected)}.xlsx")
            elif self.selection == "Aspect-General":
                os.rename("part001.xlsx", f"General_Aspect_Distribution.xlsx")
            self.master.update()
            logging.info(msg)
            msgbox.showinfo(message=msg)


    def start(
            self,
            selected: list = [],
            selected_obj: list = [],
            selected_sign: list = []
    ):
        global TBL_SS, TBL_HH_PRSNL, TBL_PSPS_PRSNL
        global TBL_HH_SYNSTRY, TBL_PSPS_SYNSTRY
        aspects = {
            ASPECTS[0]: CONJUNCTION,
            ASPECTS[1]: SEMI_SEXTILE,
            ASPECTS[2]: SEMI_SQUARE,
            ASPECTS[3]: SEXTILE,
            ASPECTS[4]: QUINTILE,
            ASPECTS[5]: SQUARE,
            ASPECTS[6]: TRINE,
            ASPECTS[7]: SESQUIQUADRATE,
            ASPECTS[8]: BIQUINTILE,
            ASPECTS[9]: QUINCUNX,
            ASPECTS[10]: OPPOSITE,
        }
        if self.selection == "Aspect-General":
            try:
                file = filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                )
                files = [i for i in open(file, "r").readlines()]
                s = len(files)
                n = time.time()
                c = 0
                num = 1
                logging.info(f"File: {os.path.split(file)[-1]}")
                logging.info(f"Number of records: {s}")
                logging.info(f"Selected: {', '.join(self.selected)}")
                logging.info(
                    f"Orb Factor: \u00b1 "
                    f"{aspects[self.selected[0].replace('-', '_')]}"
                    f"\u00b0"
                )
                logging.info(f"Mode: {', '.join(self.modes)}")
                logging.info(
                    "Calculation of 'General aspect distributions' is started."
                )
                logging.info(
                    f"Separating {len(files)} records into files..."
                )
                for i in range(len(files)):
                    if i % 400 == 0:
                        parted = files[i:i + 400]
                        part = Spreadsheet()
                        part.create_tables(
                            files=parted,
                            num=num,
                            selected=[
                                j.replace("-", "_")
                                for j in selected
                            ],
                            modes=self.modes,
                            file=file,
                            number_of_records=s
                        )
                        num += 1
                    c += 1
                    info(s=s, n=n, c=c)
                print()
                logging.info(
                    f"Completed separating {len(files)} records into files."
                )
                logging.info(f"Merging the separated files...")
                self.run()
            except:
                pass      
        elif self.selection == "Aspect-Detailed":
            try:
                file = filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                )
                files = [i for i in open(file, "r").readlines()]
                s = len(files)
                n = time.time()
                c = 0
                num = 1
                logging.info(f"File: {os.path.split(file)[-1]}")
                logging.info(f"Number of records: {s}")
                logging.info(f"Selected: {', '.join(self.selected)}")
                logging.info(
                    f"Orb Factor: \u00b1 "
                    f"{aspects[self.selected[0].replace('-', '_')]}"
                    f"\u00b0"
                )
                logging.info(f"Mode: {', '.join(self.modes)}")
                logging.info(
                    "Calculation of 'Aspect distributions of "
                    "objects' is started."
                )
                logging.info(
                    f"Separating {len(files)} records into files..."
                )
                for i in range(len(files)):
                    if i % 400 == 0:
                        parted = files[i:i + 400]
                        part = Spreadsheet()
                        part.create_tables(
                            files=parted,
                            num=num,
                            selected=[
                                j.replace("-", "_")
                                for j in selected
                            ],
                            modes=self.modes,
                            file=file,
                            number_of_records=s
                        )
                        num += 1
                    c += 1
                    info(s=s, n=n, c=c)
                print()
                logging.info(
                    f"Completed separating {len(files)} records into files."
                )
                logging.info(f"Merging the separated files...")
                self.run()
            except:
                pass
        elif self.selection == "Sign":
            try:
                ask_file = filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                )
                s = len([i for i in open(ask_file, "r").readlines()])
                logging.info(f"File: {os.path.split(ask_file)[-1]}")
                logging.info(f"Number of records: {s}")
                logging.info(
                    f"Selected Objects: {', '.join(selected_obj)}"
                )
                logging.info(f"Mode: {', '.join(self.modes)}")
                logging.info(
                    "Calculation of 'Sign positions of objects' is started."
                )
                r_object_dist_acc_to_signs_or_houses(
                    file=ask_file,
                    arg1=selected_obj[0],
                    arg2=selected_obj[1],
                    table=TBL_SS,
                    index=1,
                )
                Spreadsheet(
                    arg1=selected_obj[0],
                    arg2=selected_obj[1]
                ).w_object_dist_acc_to_signs_or_houses(
                    table=TBL_SS,
                    modes=self.modes,
                    selected_obj=selected_obj,
                    selection="Sign",
                    file=ask_file,
                    number_of_records=s
                )
                logging.info(
                    "Calculation of 'Sign positions of objects' is completed."
                )
                msgbox.showinfo(
                    message="Calculation of 'Sign positions of "
                            "objects' is completed."
                )
                TBL_SS = {
                    i: {j: 0 for j in SIGNS}
                    for i in SIGNS
                }
            except:
                pass
        elif self.selection == "Personal_House":
            try:
                ask_file = filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                )
                s = len([i for i in open(ask_file, "r").readlines()])
                logging.info(f"File: {os.path.split(ask_file)[-1]}")
                logging.info(f"Number of records: {s}")
                logging.info(
                    f"Selected Objects: {', '.join(selected_obj)}"
                )
                logging.info(
                    f"House System: {HOUSE_SYSTEMS[HSYS]}"
                )
                logging.info(f"Mode: {', '.join(self.modes)}")
                logging.info(
                    "Calculation of 'House positions of "
                    "objects (Personal)' is started."
                )
                r_object_dist_acc_to_signs_or_houses(
                    file=ask_file,
                    arg1=selected_obj[0],
                    arg2=selected_obj[1],
                    table=TBL_HH_PRSNL,
                    index=-1,
                )
                Spreadsheet(
                    arg1=selected_obj[0],
                    arg2=selected_obj[1]
                ).w_object_dist_acc_to_signs_or_houses(
                    table=TBL_HH_PRSNL,
                    modes=self.modes,
                    selected_obj=selected_obj,
                    selection="Personal_House",
                    file=ask_file,
                    number_of_records=s
                )
                logging.info(
                    "Calculation of 'House positions of "
                    "objects (Personal)' is completed."
                )
                msgbox.showinfo(
                    message="Calculation of 'House positions of "
                            "objects (Personal)' is completed."
                )
                TBL_HH_PRSNL = {
                    i: {j: 0 for j in HOUSES}
                    for i in HOUSES
                }
            except:
                pass
        elif self.selection == "Synastry_House":
            try:
                ask_file = filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                )
                s = len([i for i in open(ask_file, "r").readlines()])
                logging.info(f"File: {os.path.split(ask_file)[-1]}")
                logging.info(f"Number of records: {s}")
                logging.info(
                    f"Selected Objects: {', '.join(selected_obj)}"
                )
                logging.info(
                    f"House System: {HOUSE_SYSTEMS[HSYS]}"
                )
                logging.info(f"Mode: {', '.join(self.modes)}")
                logging.info(
                    "Calculation of 'House positions of "
                    "objects (Synastry)' is started."
                )
                r_syn_object_dist_acc_to_signs_or_houses(
                    file=ask_file,
                    arg1=selected_obj[0],
                    arg2=selected_obj[1],
                    table=TBL_HH_SYNSTRY,
                )
                Spreadsheet(
                    arg1=selected_obj[0],
                    arg2=selected_obj[1]
                ).w_object_dist_acc_to_signs_or_houses(
                    table=TBL_HH_SYNSTRY,
                    modes=self.modes,
                    selected_obj=selected_obj,
                    selection="Synastry_House",
                    file=ask_file,
                    number_of_records=s
                )
                logging.info(
                    "Calculation of 'House positions of "
                    "objects (Synastry)' is completed."
                )
                msgbox.showinfo(
                    message="Calculation of 'House positions of "
                            "objects (Synastry)' is completed."
                )
                TBL_HH_SYNSTRY = {
                    i: {j: 0 for j in HOUSES}
                    for i in HOUSES
                }
            except:
                pass
        elif self.selection == "Object_Sign":
            try:
                ask_file = filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                )
                s = len([i for i in open(ask_file, "r").readlines()])
                logging.info(f"File: {os.path.split(ask_file)[-1]}")
                logging.info(f"Number of records: {s}")
                logging.info(
                    f"Selected Objects: {', '.join(selected_obj)}"
                )
                logging.info(
                    f"Selected Signs: {', '.join(selected_sign)}"
                )
                logging.info(f"House System: {HOUSE_SYSTEMS[HSYS]}")
                logging.info(f"Mode: {', '.join(self.modes)}")
                logging.info(
                    "Calculation of 'House positions of "
                    "objects-signs (Personal)' is started."
                )
                r_object_sign_dist_acc_to_houses(
                    file=ask_file,
                    arg1=selected_obj[0],
                    arg2=selected_obj[1],
                )
                Spreadsheet(
                    arg1=selected_obj[0],
                    arg2=selected_obj[1],
                    arg3=selected_sign[0],
                    arg4=selected_sign[1]
                ).w_object_sign_dist_acc_to_houses(
                    table=TBL_PSPS_PRSNL,
                    modes=self.modes,
                    selected_obj=selected_obj,
                    selected_sign=selected_sign,
                    selection=self.selection,
                    file=ask_file,
                    number_of_records=s
                )
                logging.info(
                    "Calculation of 'House positions of "
                    "objects-signs (Personal)' is completed."
                )
                msgbox.showinfo(
                    message="Calculation of 'House positions of "
                            "objects-signs (Personal)' is completed."
                )
                TBL_PSPS_PRSNL = {
                    sign: {
                        house: {
                            _sign: {
                                _house: 0 for _house in HOUSES
                            }
                            for _sign in SIGNS
                        }
                        for house in HOUSES
                    }
                    for sign in SIGNS
                }
            except:
                pass
        elif self.selection == "Synastry_Object_Sign":
            try:
                ask_file = filedialog.askopenfilename(
                    filetypes=[("CSV File", ".csv")]
                )
                s = len([i for i in open(ask_file, "r").readlines()])
                logging.info(f"File: {os.path.split(ask_file)[-1]}")
                logging.info(f"Number of records: {s}")
                logging.info(
                    f"Selected Objects: {', '.join(selected_obj)}"
                )
                logging.info(
                    f"Selected Signs: {', '.join(selected_sign)}"
                )
                logging.info(f"House System: {HOUSE_SYSTEMS[HSYS]}")
                logging.info(f"Mode: {', '.join(self.modes)}")
                logging.info(
                    "Calculation of 'House positions of "
                    "objects-signs (Synastry)' is started."
                )
                r_syn_object_sign_dist_acc_to_houses(
                    file=ask_file,
                    arg1=selected_obj[0],
                    arg2=selected_obj[1],
                )
                Spreadsheet(
                    arg1=selected_obj[0],
                    arg2=selected_obj[1],
                    arg3=selected_sign[0],
                    arg4=selected_sign[1]
                ).w_object_sign_dist_acc_to_houses(
                    table=TBL_PSPS_SYNSTRY,
                    modes=self.modes,
                    selected_obj=selected_obj,
                    selected_sign=selected_sign,
                    selection=self.selection,
                    file=ask_file,
                    number_of_records=s
                )
                logging.info(
                    "Calculation of 'House positions "
                    "of objects-signs (Synastry)' is completed."
                )
                msgbox.showinfo(
                    message="Calculation of 'House positions "
                            "of objects-signs (Synastry)' is completed."
                )
                TBL_PSPS_SYNSTRY = {
                    sign: {
                        house: {
                            _sign: {
                                _house: 0 for _house in HOUSES
                            }
                            for _sign in SIGNS
                        }
                        for house in HOUSES
                    }
                    for sign in SIGNS
                }
            except:
                pass
        else:
            msgbox.showinfo(
                message="Please select one table to be calculated."
            )

    @staticmethod
    def check_all_buttons(check_all=None, cvar_list=[], checkbutton_list=[]):
        if check_all.get() is True:
            for var, c_button in zip(cvar_list, checkbutton_list):
                var.set(True)
                c_button.configure(variable=var)
        else:
            for var, c_button in zip(cvar_list, checkbutton_list):
                var.set(False)
                c_button.configure(variable=var)

    @staticmethod
    def check_uncheck(
            checkbuttons: dict = {},
            array: list = [],
            j: str = ""
    ):
        for i in array:
            if i != j:
                checkbuttons[i][1].set("0")
                checkbuttons[i][0].configure(variable=checkbuttons[i][1])

    def checkbutton(
            self,
            master=None,
            checkbuttons: dict = {},
            text: str = "",
            row: int = 0,
            column: int = 0,
            array: list = []
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
        checkbuttons[text][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                array,
                text
            )
        )

    def aspect_general(self):
        self.t_aspect_general = tk.Toplevel()
        self.t_aspect_general.title("General Aspect Distributions")
        self.t_aspect_general.resizable(width=False, height=False)
        main_frame = tk.Frame(master=self.t_aspect_general)
        main_frame.pack()
        left_frame = tk.Frame(
            master=main_frame, 
            bd=1, 
            relief="sunken"
        )
        left_frame.pack(side="left", fill="both")
        mid_frame = tk.Frame(
            master=main_frame, 
            bd=1, 
            relief="sunken"
        )
        mid_frame.pack(side="left", fill="both")
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
        left_cb_frame = tk.Frame(master=left_frame)
        left_cb_frame.pack()
        mid_cb_frame = tk.Frame(master=mid_frame)
        mid_cb_frame.pack()
        checkbuttons = {}
        check_all_1 = tk.BooleanVar()
        check_all_2 = tk.BooleanVar()
        check_uncheck_1 = tk.Checkbutton(
            master=left_cb_frame,
            text="Check/Uncheck All",
            variable=check_all_1)
        check_all_1.set(False)
        check_uncheck_1.grid(row=0, column=0, sticky="w")
        for i, j in enumerate(ASPECTS, 1):
            if j != "null":
                self.checkbutton(
                    master=left_cb_frame,
                    text=j.replace("_", "-"),
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
        cvar_list_1 = [i[1] for i in checkbuttons.values()][:11]
        cb_list_1 = [i[0] for i in checkbuttons.values()][:11]
        check_uncheck_1.configure(
            command=lambda: self.check_all_buttons(
                check_all_1,
                cvar_list_1,
                cb_list_1
            )
        )
        apply_button = tk.Button(
            master=self.t_aspect_general,
            text="Apply",
            command=lambda: self.select_aspects(checkbuttons)
        )
        apply_button.pack(side="bottom")
        
    def select_aspects(self, checkbuttons={}):
        self.selected = []
        for i, j in enumerate(ASPECTS):
            if "_" in j:
                j = j.replace("_", "-")
            if j != "null" and checkbuttons[j][1].get() == "1":
                self.selected.append(j)
        if checkbuttons["All Aspects"][1].get() == "1":
            self.selected.append("all")
        if len(self.selected) < 1:
            msgbox.showinfo(
                message="Please select at least one aspect."
            )
        else:
            self.t_aspect_general.destroy()
            self.t_aspect_general = None
            self.selected_obj = []
            self.selected_sign = []
            self.selection = "Aspect-General"

    def select_tables(self, checkbuttons: dict = {}):
        self.selected = []
        for i, j in enumerate(ASPECTS):
            if "_" in j:
                j = j.replace("_", "-")
            if j != "null" and checkbuttons[j][1].get() == "1":
                self.selected.append(j)
        for i in OBJECTS:
            if checkbuttons[i][1].get() == "1":
                self.selected.append(i)
        if len(self.selected) != 2:
            msgbox.showinfo(
                message="Please select one aspect and one object."
            )
        else:
            self.t_aspect_dist_acc_to_objects.destroy()
            self.t_aspect_dist_acc_to_objects = None
            self.selected_obj = []
            self.selected_sign = []
            self.selection = "Aspect-Detailed"

    def aspect_dist_acc_to_objects(self):
        self.t_aspect_dist_acc_to_objects = tk.Toplevel()
        self.t_aspect_dist_acc_to_objects.title(
            "Aspect distributions of objects"
        )
        self.t_aspect_dist_acc_to_objects.geometry("400x420")
        self.t_aspect_dist_acc_to_objects.resizable(
            width=False, height=False
        )
        main_frame = tk.Frame(master=self.t_aspect_dist_acc_to_objects)
        main_frame.pack()
        left_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken",
        )
        left_frame.pack(side="left", fill="both")
        right_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        right_frame.pack(side="left", fill="both")
        left_frame_label = tk.Label(
            master=left_frame,
            text="Select Aspect Types",
            fg="red"
        )
        left_frame_label.pack()
        right_frame_label = tk.Label(
            master=right_frame,
            text="Select Objects",
            fg="red"
        )
        right_frame_label.pack()
        left_cb_frame = tk.Frame(master=left_frame)
        left_cb_frame.pack()
        right_cb_frame = tk.Frame(master=right_frame)
        right_cb_frame.pack()
        checkbuttons = {}
        for i, j in enumerate(ASPECTS, 1):
            if j != "null":
                self.checkbutton(
                    master=left_cb_frame,
                    text=j.replace("_", "-"),
                    row=i,
                    column=0,
                    checkbuttons=checkbuttons,
                    array=[
                        asp.replace("_", "-")
                        for asp in ASPECTS if asp != "null"
                    ]
                )
        for i, j in enumerate(OBJECTS, 1):
            self.checkbutton(
                master=right_cb_frame,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons,
                array=OBJECTS
            )
        fill_left = tk.Frame(left_cb_frame, height=46)
        fill_left.grid(row=12, column=0)
        apply_button = tk.Button(
            master=self.t_aspect_dist_acc_to_objects,
            text="Apply",
            command=lambda: self.select_tables(checkbuttons)
        )
        apply_button.pack(side="bottom")

    def select_objects(
            self,
            checkbuttons1: dict = {},
            checkbuttons2: dict = {}
    ):
        self.selected_obj = []
        for i in OBJECTS:
            if checkbuttons1[i][1].get() == "1":
                self.selected_obj.append(i)
        for i in OBJECTS:
            if checkbuttons2[i][1].get() == "1":
                self.selected_obj.append(i)
        if len(self.selected_obj) != 2:
            msgbox.showinfo(
                message="Please select at least two objects."
            )
        else:
            self.t_object_dist_acc_to_signs.destroy()
            self.t_object_dist_acc_to_signs = None
            self.selected = []
            self.selected_sign = []

    def object_dist_acc_to_signs(self):
        self.selection = "Sign"
        self.t_object_dist_acc_to_signs = tk.Toplevel()
        self.t_object_dist_acc_to_signs.title(
            "Sign positions of objects"
        )
        self.t_object_dist_acc_to_signs.geometry("400x420")
        self.t_object_dist_acc_to_signs.resizable(
            width=False, height=False
        )
        main_frame = tk.Frame(master=self.t_object_dist_acc_to_signs)
        main_frame.pack()
        left_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        left_frame.pack(side="left")
        label_left = tk.Label(
            master=left_frame,
            text="Select an object\nfor males",
            fg="red"
        )
        label_left.pack()
        left_bottom = tk.Frame(master=left_frame)
        left_bottom.pack()
        right_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        right_frame.pack(side="left")
        label_right = tk.Label(
            master=right_frame,
            text="Select an object\nfor females",
            fg="red"
        )
        label_right.pack()
        right_bottom = tk.Frame(master=right_frame)
        right_bottom.pack()
        checkbuttons1 = {}
        checkbuttons2 = {}
        for i, j in enumerate(OBJECTS, 1):
            self.checkbutton(
                master=left_bottom,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons1,
                array=OBJECTS
            )
            self.checkbutton(
                master=right_bottom,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons2,
                array=OBJECTS
            )
        apply_button = tk.Button(
            master=self.t_object_dist_acc_to_signs,
            text="Apply",
            command=lambda: self.select_objects(
                checkbuttons1,
                checkbuttons2
            )
        )
        apply_button.pack(side="bottom")

    def select_houses(
            self,
            checkbuttons1: dict = {},
            checkbuttons2: dict = {}
    ):
        self.selected_obj = []
        for i in OBJECTS:
            if checkbuttons1[i][1].get() == "1":
                self.selected_obj.append(i)
        for i in OBJECTS:
            if checkbuttons2[i][1].get() == "1":
                self.selected_obj.append(i)
        if len(self.selected_obj) != 2:
            msgbox.showinfo(
                message="Please select at least two objects."
            )
        else:
            self.t_object_dist_acc_to_houses.destroy()
            self.t_object_dist_acc_to_houses = None
            self.selected = []
            self.selected_sign = []

    def object_dist_acc_to_houses(self, title, selection):
        self.selection = selection
        self.t_object_dist_acc_to_houses = tk.Toplevel()
        self.t_object_dist_acc_to_houses.title(title)
        self.t_object_dist_acc_to_houses.geometry("400x420")
        self.t_object_dist_acc_to_houses.resizable(
            width=False, height=False
        )
        main_frame = tk.Frame(master=self.t_object_dist_acc_to_houses)
        main_frame.pack()
        left_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        left_frame.pack(side="left")
        label_left = tk.Label(
            master=left_frame,
            text="Select an object\nfor males",
            fg="red"
        )
        label_left.pack()
        left_bottom = tk.Frame(master=left_frame)
        left_bottom.pack()
        right_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        right_frame.pack(side="left")
        label_right = tk.Label(
            master=right_frame,
            text="Select an object\nfor females",
            fg="red"
        )
        label_right.pack()
        right_bottom = tk.Frame(master=right_frame)
        right_bottom.pack()
        checkbuttons1 = {}
        checkbuttons2 = {}
        for i, j in enumerate(OBJECTS, 1):
            self.checkbutton(
                master=left_bottom,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons1,
                array=OBJECTS
            )
            self.checkbutton(
                master=right_bottom,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons2,
                array=OBJECTS
            )
        apply_button = tk.Button(
            master=self.t_object_dist_acc_to_houses,
            text="Apply",
            command=lambda: self.select_houses(
                checkbuttons1,
                checkbuttons2
            )
        )
        apply_button.pack(side="bottom")

    def select_objects_signs(
            self,
            checkbuttons1: dict = {},
            checkbuttons2: dict = {},
            checkbuttons3: dict = {},
            checkbuttons4: dict = {}
    ):
        self.selected_obj = []
        self.selected_sign = []
        for i in OBJECTS:
            if checkbuttons1[i][1].get() == "1":
                self.selected_obj.append(i)
        for i in OBJECTS:
            if checkbuttons2[i][1].get() == "1":
                self.selected_obj.append(i)
        for i in SIGNS:
            if checkbuttons3[i][1].get() == "1":
                self.selected_sign.append(i)
        for i in SIGNS:
            if checkbuttons4[i][1].get() == "1":
                self.selected_sign.append(i)
        if len(self.selected_obj) != 2 and len(self.selected_sign) != 2:
            msgbox.showinfo(
                message="Please select at least two objects and "
                        "two signs."
            )
        else:
            self.t_object_sign_dist_acc_to_houses.destroy()
            self.t_object_sign_dist_acc_to_houses = None
            self.selected = []

    def object_sign_dist_acc_to_houses(self, title, selection):
        self.selection = selection
        self.t_object_sign_dist_acc_to_houses = tk.Toplevel()
        self.t_object_sign_dist_acc_to_houses.title(title)
        self.t_object_sign_dist_acc_to_houses.geometry("400x420")
        self.t_object_sign_dist_acc_to_houses.resizable(
            width=False, height=False
        )
        main_frame = tk.Frame(master=self.t_object_sign_dist_acc_to_houses)
        main_frame.pack()
        left_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        left_frame.pack(side="left")
        label_left = tk.Label(
            master=left_frame,
            text="Select an object and a sign\nfor males",
            fg="red"
        )
        label_left.pack()
        left_bottom = tk.Frame(master=left_frame)
        left_bottom.pack()
        right_frame = tk.Frame(
            master=main_frame,
            bd=1,
            relief="sunken"
        )
        right_frame.pack(side="left")
        label_right = tk.Label(
            master=right_frame,
            text="Select an object and a sign \nfor females",
            fg="red"
        )
        label_right.pack()
        right_bottom = tk.Frame(master=right_frame)
        right_bottom.pack()
        checkbuttons1 = {}
        checkbuttons2 = {}
        checkbuttons3 = {}
        checkbuttons4 = {}
        for i, j in enumerate(OBJECTS, 1):
            self.checkbutton(
                master=left_bottom,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons1,
                array=OBJECTS
            )
            self.checkbutton(
                master=right_bottom,
                text=j,
                row=i,
                column=0,
                checkbuttons=checkbuttons2,
                array=OBJECTS
            )
        for i, j in enumerate(SIGNS, 1):
            self.checkbutton(
                master=left_bottom,
                text=j,
                row=i,
                column=1,
                checkbuttons=checkbuttons3,
                array=SIGNS
            )
            self.checkbutton(
                master=right_bottom,
                text=j,
                row=i,
                column=1,
                checkbuttons=checkbuttons4,
                array=SIGNS
            )
        apply_button = tk.Button(
            master=self.t_object_sign_dist_acc_to_houses,
            text="Apply",
            command=lambda: self.select_objects_signs(
                checkbuttons1,
                checkbuttons2,
                checkbuttons3,
                checkbuttons4
            )
        )
        apply_button.pack(side="bottom")

    def change_mode(self, variables: list = []):
        self.modes = [i.get() for i in variables]
        self.t_mode.destroy()
        self.t_mode = None

    def mode(self):
        self.t_mode = tk.Toplevel()
        self.t_mode.title("Mode")
        self.t_mode.resizable(width=False, height=False)
        self.modes = []
        main_frame = tk.Frame(master=self.t_mode)
        main_frame.pack()
        variables = []
        for i, j in enumerate(("Males", "Females")):
            frame = tk.Frame(master=main_frame, relief="sunken", bd=1)
            frame.grid(row=0, column=i)
            label = tk.Label(
                master=frame,
                text=j,
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
            master=self.t_mode,
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
        self.t_orb = tk.Toplevel()
        self.t_orb.title("Orb Factor")
        self.t_orb.resizable(width=False, height=False)
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
            aspect_label = tk.Label(master=self.t_orb, text=f"{j}")
            aspect_label.grid(row=i, column=0, sticky="w")
            equal_to = tk.Label(master=self.t_orb, text="=")
            equal_to.grid(row=i, column=1, sticky="e")
            orb_entry = tk.Entry(master=self.t_orb, width=5)
            orb_entry.grid(row=i, column=2)
            orb_entry.insert(0, default_orbs[i])
            orb_entries.append(orb_entry)
        apply_button = tk.Button(
            master=self.t_orb,
            text="Apply",
            command=lambda: self.change_orb_factors(
                orb_entries=orb_entries
            )
        )
        apply_button.grid(row=11, column=0, columnspan=3)

    def change_hsys(self, checkbuttons, _house_systems_):
        global HSYS
        if checkbuttons[_house_systems_[0]][1].get() == "1":
            HSYS = "P"
        elif checkbuttons[_house_systems_[1]][1].get() == "1":
            HSYS = "K"
        elif checkbuttons[_house_systems_[2]][1].get() == "1":
            HSYS = "O"
        elif checkbuttons[_house_systems_[3]][1].get() == "1":
            HSYS = "R"
        elif checkbuttons[_house_systems_[4]][1].get() == "1":
            HSYS = "C"
        elif checkbuttons[_house_systems_[5]][1].get() == "1":
            HSYS = "E"
        elif checkbuttons[_house_systems_[6]][1].get() == "1":
            HSYS = "W"
        self.t_hsys.destroy()
        self.t_hsys = None

    def create_hsys_checkbuttons(self):
        self.t_hsys = tk.Toplevel()
        self.t_hsys.title("House System")
        self.t_hsys.geometry("200x200")
        self.t_hsys.resizable(width=False, height=False)
        hsys_frame = tk.Frame(master=self.t_hsys)
        hsys_frame.pack(side="top")
        button_frame = tk.Frame(master=self.t_hsys)
        button_frame.pack(side="bottom")
        _house_systems_ = [values for keys, values in HOUSE_SYSTEMS.items()]
        checkbuttons = {}
        for i, j in enumerate(_house_systems_):
            _var_ = tk.StringVar()
            if j == HOUSE_SYSTEMS[HSYS]:
                _var_.set(value="1")
            else:
                _var_.set(value="0")
            _checkbutton_ = tk.Checkbutton(
                master=hsys_frame,
                text=j,
                variable=_var_)
            _checkbutton_.grid(row=i, column=0, sticky="w")
            checkbuttons[j] = [_checkbutton_, _var_]
        checkbuttons["Placidus"][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                _house_systems_,
                "Placidus"
            )
        )
        checkbuttons["Koch"][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                _house_systems_,
                "Koch"
            )
        )
        checkbuttons["Porphyrius"][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                _house_systems_,
                "Porphyrius"
            )
        )
        checkbuttons["Regiomontanus"][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                _house_systems_,
                "Regiomontanus"
            )
        )
        checkbuttons["Campanus"][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                _house_systems_,
                "Campanus"
            )
        )
        checkbuttons["Equal"][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                _house_systems_,
                "Equal"
            )
        )
        checkbuttons["Whole Signs"][0].configure(
            command=lambda: self.check_uncheck(
                checkbuttons,
                _house_systems_,
                "Whole Signs"
            )
        )
        apply_button = tk.Button(
            master=button_frame,
            text="Apply",
            command=lambda: self.change_hsys(
                checkbuttons,
                _house_systems_
            )
        )
        apply_button.pack()

    def change_orb_factors(self, orb_entries: list = []):
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
        self.t_orb.destroy()
        self.t_orb = None

    @staticmethod
    def update_script():
        url_1 = "https://raw.githubusercontent.com/dildeolupbiten/" \
                "SynastryParser/master/SynastryParser.py"
        url_2 = "https://raw.githubusercontent.com/dildeolupbiten/" \
                "SynastryParser/master/README.md"
        data_1 = urllib.request.urlopen(
            url=url_1,
            context=ssl.SSLContext(ssl.PROTOCOL_SSLv23))
        data_2 = urllib.request.urlopen(
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
    def callback(url):
        webbrowser.open_new(url)

    def about(self):
        tl = tk.Toplevel()
        tl.title("About SynastryParser")
        tl.resizable(height=False, width=False)
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
        thanks_to, _thanks_to = "Thanks To:", "Flavia Alonso"
        cura = "C.U.R.A."
        blank, _thanks_to_ = " " * len("Thanks To:"), cura
        contact, _contact = "Contact:", "tckutlu@gmail.com"
        github, _github = "GitHub:", \
            "https://github.com/dildeolupbiten/SynastryParser"
        tframe1 = tk.Frame(master=tl, bd="2", relief="groove")
        tframe1.pack(fill="both")
        tframe2 = tk.Frame(master=tl)
        tframe2.pack(fill="both")
        tlabel_title = tk.Label(
            master=tframe1, text=name, font="Arial 25"
        )
        tlabel_title.pack()
        for i, j in enumerate([
            version, build_date, update_date, thanks_to,
            blank, developed_by, contact, github
        ]):
            tlabel_info_1 = tk.Label(master=tframe2, text=j,
                                     font="Arial 11", fg="red")
            tlabel_info_1.grid(row=i, column=0, sticky="w")
        for i, j in enumerate((
                _version, _build_date, _update_date, _thanks_to,
                _thanks_to_, _developed_by, _contact, _github
        )):
            if j == _github:
                tlabel_info_2 = tk.Label(master=tframe2, text=j,
                                         font="Arial 11", fg="blue",
                                         cursor="hand2")
                url1 = "https://github.com/dildeolupbiten/SynastryParser"
                tlabel_info_2.bind(
                    "<Button-1>",
                    lambda event: self.callback(url1))
            elif j == _contact:
                tlabel_info_2 = tk.Label(
                    master=tframe2, 
                    text=j,
                    font="Arial 11", 
                    fg="blue",
                    cursor="hand2"
                )
                url2 = "mailto://tckutlu@gmail.com"
                tlabel_info_2.bind(
                    "<Button-1>",
                    lambda event: self.callback(url2))
            elif j == _thanks_to_:
                tlabel_info_2 = tk.Label(
                    master=tframe2, 
                    text=j,
                    font="Arial 11", 
                    fg="blue",
                    cursor="hand2"
                )
                url3 = "http://cura.free.fr/"
                tlabel_info_2.bind(
                    "<Button-1>",
                    lambda event: self.callback(url3))
            else:
                tlabel_info_2 = tk.Label(
                    master=tframe2, 
                    text=j,
                    font="Arial 11"
                )
            tlabel_info_2.grid(row=i, column=1, sticky="w")


def main():
    root = tk.Tk()
    root.title("SynastryParser")
    root.resizable(height=False, width=False)
    app = App(master=root)
    threading.Thread(target=app.mainloop).run()


if __name__ == "__main__":
    main()
