import requests
import json
import openpyxl
import datetime
import time
import sys
import os

class HeroRec():
    def __init__(self):
        with open(os.path.dirname(os.path.abspath(sys.argv[0])) + "/config.json", mode="r") as jf:
            df = json.load(jf)
            self.levellist = df["levellist"]
            self.levelidlist = df["levelidlist"]
            self.weaponid = df["weaponid"]
            self.weaponlist = df["weaponlist"]
            self.weaponiddict = df["weaponiddict"]
            self.invalidrunlist = df["invalidrunlist"] # not used

    def opensheet(self):
        try:
            wbpath = os.path.dirname(os.path.abspath(sys.argv[0])) + "/HeromodeILrecords.xlsx"
            self.wb = openpyxl.load_workbook(wbpath)
            self.ws = self.wb.worksheets[0]
        except FileNotFoundError: # file is not exist
            self.wb = openpyxl.Workbook() # create new file
            self.ws = self.wb.worksheets[0]
            self.sheetsetup()

    def cell(self, row, column):
        return self.ws.cell(row+2, column+2) # set (2, 2) to datum

    def sheetsetup(self):
        # level
        tmp = [3, 6, 6, 6, 6]
        n = 0
        level = 1
        for i in range(5):
            for _ in range(tmp[i]):
                self.cell(n, -1).value = f"{level:0>2}"
                self.cell(n, 11).value = f"{level:0>2}"
                n += 1
                level += 1
            self.cell(n, -1).value = f"Boss{i+1}"
            self.cell(n, 11).value = f"Boss{i+1}"
            n += 1
        # weapon
        for m in range(9):
            self.cell(-1, m).value = self.weaponlist[m]
        self.cell(-1, 12).value = "WR"
        self.cell(-1, 13).value = "weapon"
        # format
        for n in range(32):
            for m in range(9):
                self.cell(n, m).number_format = "[mm]:ss"
            self.cell(n, 11).number_format = "[mm]:ss"

    def getrec(self, levelnum):
        levelurl = f"https://www.speedrun.com/api/v1/runs?status=verified&max=200&level={self.levelidlist[levelnum]}" # first page url
        leveldata = []
        while True:
            levelinfo = requests.get(levelurl).json()
            time.sleep(5)
            leveldata += levelinfo["data"]
            # search for next page
            levellinks = levelinfo["pagination"]["links"]
            for link in levellinks:
                if link["rel"] == "next":
                    levelurl = link["uri"] # get next page url
                    break
            else:
                break # break while loop

        weaponrec = [
            [
                rundata["times"]["primary_t"]
                for rundata in leveldata
                if not rundata["id"] in self.invalidrunlist
                and self.weaponiddict[rundata["values"][self.weaponid]] == n
            ]
            for n in range(9)
        ]
        weaponrec = [
            min(weaponrec[n])
            if weaponrec[n]
            else False
            for n in range(9)
        ]
        # 次と同値
        """
        weaponrec = [False]*9
        # check all run data
        for rundata in leveldata:
            if not rundata["id"] in self.invalidrunlist:
                runweaponid = rundata["values"][self.weaponid]
                weaponnum = self.weaponiddict[runweaponid]
                runrec = rundata["times"]["primary_t"]
                currentrec = weaponrec[weaponnum]
                if not currentrec or currentrec > runrec:
                    weaponrec[weaponnum] = runrec # renew WR
        """

        return weaponrec

    def mainfunc(self):
        self.opensheet()
        WRlist = [None]*32
        for n in range(32):
            print(f"get records in {self.levellist[n]}")
            WRlist[n] = self.getrec(n)

        # convert to datetime.time
        for n in range(32):
            for m in range(9):
                if WRlist[n][m]: # has value
                    WRlist[n][m] = datetime.time(
                        WRlist[n][m]//3600,
                        WRlist[n][m]%3600//60,
                        WRlist[n][m]%60
                    )

        # find new record
        for n in range(32):
            for m in range(9):
                if WRlist[n][m]:
                    if self.cell(n, m).value:
                        if WRlist[n][m] < self.cell(n, m).value:
                            print(f"New record! {self.levellist[n]} {self.weaponlist[m]}, {WRlist[n][m].hour*60+WRlist[n][m].minute:0>2}:{WRlist[n][m].second:0>2}")
                    else:
                        print(f"New record! {self.levellist[n]} {self.weaponlist[m]}, {WRlist[n][m].hour*60+WRlist[n][m].minute:0>2}:{WRlist[n][m].second:0>2}")

        # renew records in Excel file
        for n in range(32):
            for m in range(9):
                if WRlist[n][m]:
                    self.cell(n, m).value = WRlist[n][m]
                else:
                    self.cell(n, m).value = None
            if any(WRlist[n]): # any weapon records
                tmp = [WRlist[n][m] for m in range(9) if WRlist[n][m]] #exclude False
                self.cell(n, 12).value = min(tmp)
                self.cell(n, 13).value = self.weaponlist[tmp.index(min(tmp))]

        self.cell(-1, -1).value = datetime.datetime.now().strftime("%Y/%m/%d")
        self.wb.save("HeromodeILrecords.xlsx")


if __name__ == "__main__":
    a = HeroRec()
    a.mainfunc()
