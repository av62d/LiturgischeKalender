from asyncio.proactor_events import _ProactorDuplexPipeTransport
from asyncio.windows_events import NULL
from enum import Enum
from tabnanny import verbose
from typing import OrderedDict

import locale

from dateutil.rrule import *
from dateutil.easter import *
from datetime import date
from datetime import datetime
from datetime import *; from dateutil.relativedelta import *

from dataclasses import dataclass

import xlsxwriter

class ColorType():
    GREEN='groen'
    WHITE='wit'
    PURPLE='paars'
    ROSA='roze'
    RED='rood'

class ColorChangeType(Enum):
    UNTIL_INC = 0
    UNTIL_EXC = 1
    AFTER_INC = 2
    AFTER_EXC = 3
    SINGLEDAY = 4

class Ranks(Enum):
    WEEKDAY = 0
    COMMEMORATION = 1
    OPTIONAL = 2
    MEMORIAL = 3
    FEAST = 4
    SUNDAY = 5
    LORD = 6
    ASHWED = 7
    HOLYWEEK = 8
    TRIDUUM = 9
    SOLEMNITY =10

# class Seasons(Enum):
#     ORDINARY = 0
#     ADVENT = 1
#     CHRISTMAS = 2
#     LENT = 3
#     EASTER = 4
#     SUMMER = 5
#     FALL = 6

@dataclass
class LiturgicalDay():
    dt: datetime
    color: ColorType
    descr: str = ""

@dataclass
class ColorChange():
    cc_day: LiturgicalDay
    cc_from_color: ColorType
    cc_to_color: ColorType
    cc_type: ColorChangeType
    cc_descr: str = ""




class LiturgicalCalendar():
    def __init__(self, year=None, verbose=False):
        locale.setlocale(locale.LC_ALL, 'nl_NL')
        if (year):
            self.year = year
        else:
            self.year = date.today().year

        self.verbose = verbose
        self.dayList = []
        self.colorChangeList = []
        self.fd_php = NULL  # File descriptor php file
        self.fd_txt = NULL  # File descriptor text file
        self.generateCalender()

    def addDay(self, dt, color=ColorType.GREEN, descr=""):
        d = LiturgicalDay(dt, color, descr)
        self.dayList.append(d)
        return d

    def addColorChange(self, ld: LiturgicalDay, cc_from_color: ColorType, cc_to_color: ColorType, cc_type: ColorChange, cc_descr = ""):
        cc = ColorChange(ld, cc_from_color, cc_to_color, cc_type, cc_descr)
        self.colorChangeList.append(cc)


    # date to datetime
    #def _dtodt(self, d, h=0, m=0):
        #return datetime(d.year, d.month, d.day, h, m)

    def setDates(self):
        print("Obsoleted setDates called")
    

    def setDtTime(self,dt,hour=0,minute=0):
        newDt = datetime(dt.year, dt.month, dt.day, hour, minute)
        return newDt


    def generateCalender(self):
        # Compute the whole calendar for given year

        # print("generating cal for {}".format(self.year))

        nieuwjaar = datetime(self.year, 1, 1, 10, 0)
        oudjaar = datetime(self.year, 12, 31, 19,30)
       

        self.setOfSundays = rrule(freq=DAILY,
                                  byweekday=(SU),
                                  dtstart=nieuwjaar,
                                  until=oudjaar)

        self.epifanie = datetime(self.year, 1, 6)

        # calculate easter and all derived dates
        e = easter(self.year, method=3)
        self.pasen = datetime(e.year,e.month,e.day, 10, 0)
        self.aswoensdag = self.pasen + relativedelta(days=-46)
        self.palmzondag = self.pasen + relativedelta(weeks=-1)
        self.wittedonderdag = self.setDtTime(self.pasen + relativedelta(days=-3), 19,30)
        self.goedevrijdag = self.setDtTime(self.pasen + relativedelta(days=-2), 19, 30)
        self.paaswake = self.setDtTime(self.pasen + relativedelta(days=-1), 21,30)
        self.hemelvaart = self.pasen + relativedelta(days=39)
        self.pinksteren = self.pasen + relativedelta(weeks=7)
        self.trinitatis = self.pinksteren + relativedelta(weeks=1)
        self.beginZomer = self.trinitatis + relativedelta(days=+1,weeks=2)
        # Christmas is fixed, dates of End of Summer, Fall and Advent are derived from Sunday before Christmas (4th Advent)
        self.kerstnacht = datetime(self.year,12,24,21,30)
        self.kerstmis = datetime(self.year,12,25,10,0)
        # last Sunday before Christmas is 4th Advent, 1st Advent and EndOfSummer are derived from that
        self.vierdeAdvent = self.kerstmis + relativedelta(days=-1,weekday=SU(-1))
        self.eersteAdvent = self.vierdeAdvent + relativedelta(weeks=-3)
        self.eindeZomer = self.eersteAdvent + relativedelta(weeks=-11)


        sundaysOfChristmasJanuary = self.setOfSundays.between(nieuwjaar,self.epifanie,inc=True)
        sundaysOfEpifany = self.setOfSundays.between(self.epifanie,self.aswoensdag,inc=True)
        sundaysOfLent = self.setOfSundays.between(self.aswoensdag,self.palmzondag,inc=True)
        sundaysOfEaster = self.setOfSundays.between(self.pasen,self.pinksteren,inc=False)
        sundaysOfTrinitatis = self.setOfSundays.between(self.trinitatis,self.beginZomer,inc=False)
        sundaysOfSummer = self.setOfSundays.between(self.beginZomer,self.eindeZomer,inc=True)
        sundaysOfFall = self.setOfSundays.between(self.eindeZomer,self.eersteAdvent,inc=False)
        sundaysOfAdvent = self.setOfSundays.between(self.eersteAdvent,self.vierdeAdvent,inc=True)
        sundaysOfChristmasDecember = self.setOfSundays.between(self.kerstmis,oudjaar,inc=True)

        self.addDay(nieuwjaar, ColorType.WHITE, "Nieuwjaar")

        i = 0
        for s in sundaysOfChristmasJanuary:
            i += 1
            # print("{} {}e zondag van kerst".format(s, i))
            self.addDay(s, ColorType.WHITE, "{}e zondag van kerst".format(i))

  
        d = self.addDay(self.epifanie, ColorType.WHITE, "Epifanie")
        self.addColorChange(d, ColorType.WHITE,ColorType.GREEN,ColorChangeType.UNTIL_INC, "Einde kersttijd")

        # self.addColorChange(self.epifanie, ColorType.WHITE, ColorType.GREEN,ColorChangeType.UNTIL_INC)

        i = 0
        for s in sundaysOfEpifany:
            i += 1
            # print("{} {}e zondag na epifanie".format(s, i))
            self.addDay(s, ColorType.GREEN, "{}e zondag na epifanie".format(i))

       
        d = self.addDay(self.aswoensdag, ColorType.PURPLE, "Aswoensdag")
        self.addColorChange(d, ColorType.GREEN, ColorType.PURPLE,ColorChangeType.AFTER_INC, "Begin 40dagentijd")

        i = 0
        for s in sundaysOfLent:
            i += 1
            if (i==4):
                c = ColorType.ROSA
            else:
                c = ColorType.PURPLE
            self.addDay(s, c, "{}e zondag van de veertigdagentijd".format(i))

        # print ("{} palmzondag".format(self.palmzondag))
        self.addDay(self.palmzondag, ColorType.PURPLE,  "Palmzondag")


        d = self.addDay(self.wittedonderdag, ColorType.WHITE, "Witte donderdag")
        self.addColorChange(d, ColorType.PURPLE, ColorType.WHITE,ColorChangeType.SINGLEDAY)

        d = self.addDay(self.goedevrijdag, ColorType.RED, "Goede Vrijdag")
        self.addColorChange(d, ColorType.PURPLE, ColorType.RED,ColorChangeType.SINGLEDAY)

        d = self.addDay(self.paaswake, ColorType.WHITE, "Paaswake")
        self.addColorChange(d, ColorType.PURPLE, ColorType.WHITE,ColorChangeType.AFTER_INC, "Begin Paastijd")

        # print ("{} pasen".format(self.pasen))
        self.addDay(self.pasen, ColorType.WHITE, "Pasen")

        i = 0
        for s in sundaysOfEaster:
            i += 1
            # print("{} {}e zondag van pasen".format(s, i))
            self.addDay(s, ColorType.WHITE, "{}e zondag van Pasen".format(i))

        # print ("{} hemelvaart".format(self.hemelvaart))
        self.addDay(self.hemelvaart, ColorType.WHITE,  "Hemelvaart")

        # print ("{} pinksteren".format(self.pinksteren))
        d = self.addDay(self.pinksteren, ColorType.RED,  "Pinksteren")

        self.addColorChange(d, ColorType.WHITE, ColorType.RED, ColorChangeType.SINGLEDAY)
        self.addColorChange(d, ColorType.WHITE, ColorType.GREEN, ColorChangeType.AFTER_EXC, "Begin Zomer")


        d = self.addDay(self.trinitatis, ColorType.WHITE,  "Trinitatis")
        self.addColorChange(d, ColorType.GREEN,  ColorType.WHITE, ColorChangeType.SINGLEDAY, "Trinitatis(drievuldigheid)")

        i = 0
        for s in sundaysOfTrinitatis:
            i += 1
            n = self.setDtTime(s, 10, 0)
            self.addDay(s, ColorType.GREEN, "{}e zondag na Trinitatis".format(i))

        i = 0
        for s in sundaysOfSummer:
            i += 1
            n = self.setDtTime(s, 10, 0)
            self.addDay(s, ColorType.GREEN, "{}e zondag van de Zomer".format(i))

        i = 0
        for s in sundaysOfFall:
            i += 1
            # print("{} {}e zondag van de herfst".format(s, i))
            self.addDay(s, ColorType.GREEN,  "{}e zondag van de Herfst".format(i))

        i = 0
        for s in sundaysOfAdvent:
            i += 1
            if (i==3):
                c = ColorType.ROSA
            else:
                c = ColorType.PURPLE
            # print("{} {}e zondag van de Advent".format(s, i))
            d = self.addDay(s, c,  "{}e zondag van de Advent".format(i))
            if (i == 1):
                self.addColorChange(d, ColorType.GREEN, ColorType.PURPLE, ColorChangeType.AFTER_INC, "Eerste Advent")
        
        d =  self.addDay(self.kerstnacht, ColorType.WHITE,  "Kerstnacht")
        self.addColorChange(d, ColorType.PURPLE, ColorType.WHITE, ColorChangeType.AFTER_INC, "Begin Kersttijd")
        # print ("{} kerst".format(self.kerstmis))
        self.addDay(self.kerstmis, ColorType.WHITE,  "Kerstmis")
       

        i = 0
        for s in sundaysOfChristmasDecember:
            i += 1
            self.addDay(s, ColorType.WHITE,  "zondag van het kerstoctaaf".format(i))
        
                
        self.addDay(oudjaar, ColorType.WHITE,  "Oudjaar")

        # sort all dates
        self.dayList.sort(key=lambda x: (x.dt.month, x.dt.day))

    def printCal(self):
        for d in self.dayList:
            print (d.dt.strftime("%a %e %b, %H:%M"), end=" ")
            print (d.color, d.descr)


    def printTXT(self, msg):
        if (self.fd_txt):
            self.fd_txt.write(msg)
            self.fd_txt.write('\n')

    def genTXTLiturgicalCalendar(self):
        self.fd_txt_name = "litcal_{}.txt".format(self.year)
        print ("Writing file {} for year {}".format(self.fd_txt_name,self.year))
        self.fd_txt = open (self.fd_txt_name, 'w')
        self.printTXT ("Jaar {}".format(self.year))
        for d in self.dayList:
            self.printTXT ("{}, {}, {}". format(d.dt.strftime("%a %e %b %y, %H:%M"), d.color, d.descr))

        self.fd_txt.close()

    def genXLSXLiturgicalCalendar(self):
        self.fd_xlsx_name = "litcal_{}.xlsx".format(self.year)
        print ("Writing file {} for year {}".format(self.fd_xlsx_name,self.year))
        workbook = xlsxwriter.Workbook(self.fd_xlsx_name)
        worksheet = workbook.add_worksheet()
        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})

        # Add a number format for cells with money.
        money = workbook.add_format({'num_format': '$#,##0'})
        # Start from the first cell. Rows and columns are zero indexed.
        row = 1
        worksheet.write('A1', 'Datum', bold)
        worksheet.write('B1', 'Kleur', bold)
        worksheet.write('C1', 'Beschrijving', bold)
        date_format = workbook.add_format({'num_format': 'ddd d mmm yy HH:MM'})
        col_red  = workbook.add_format({'bg_color': 'red'})
        col_green  = workbook.add_format({'bg_color': 'green'})
        col_white  = workbook.add_format({'bg_color': 'white'})
        col_purple  = workbook.add_format({'bg_color': 'purple'})
        col_rosa  = workbook.add_format({'bg_color': 'rosa'})
        for d in self.dayList:
            date_str = d.dt.strftime('%Y-%m-%d %H:%M')
            fmt = col_green
            if (d.color == 'rood'):
                fmt = col_red
            elif (d.color == 'wit'):
                fmt =  col_white
            elif (d.color == 'paars'):
                fmt = col_purple
            # date_format.set_bg_color(col_green)
            # str_format.set_bg_color(bg_color)
            worksheet.write_datetime(row, 0, d.dt, date_format)
            worksheet.write_string(row, 1, d.color, fmt)
            worksheet.write_string(row, 2, d.descr, fmt)

            row += 1

        worksheet.add_table('A1:C{}'.format(row), {
            'style': 'Table Style Light 11',
            'columns': [{'header': 'Datum'},
                        {'header': 'Kleur'},
                        {'header': 'Beschrijving'}
                        ]})

        workbook.close()

    def printPHP(self, indent, msg):
        if (self.fd_php):
            for i in range(0,indent):
                self.fd_php.write("   ")
            self.fd_php.write(msg)
            self.fd_php.write('\n')

    def genPHPLiturgicalCalendar(self):
        self.fd_php_name = "litcal_{}.php".format(self.year)
        print ("Writing file {} for year {}".format(self.fd_php_name,self.year))
        self.fd_php = open (self.fd_php_name, 'w')
        self.printPHP (0, "# Year {}".format(self.year))
        header="""
<?php
// BEGIN color function
function set_liturgical_color(&$light, &$dark, &$font)
{
    $jaar = date("y") + 2000;
    $dag = date("d");
    $maand = date("m");
    // kies 1 van vier kleuren: groen, goud, paars of rood
        """

        self.printPHP(0, header)
        self.printPHP(1, "$kleur = '{}';".format(ColorType.GREEN))

        self.printPHP(1, "if ($jaar == {}) {{".format(self.year))
        self.printPHP(2, "switch ($maand) {")
        cmonth = -1
        cday = -1
        ccolor = ColorType.GREEN

        for cc in self.colorChangeList:
            ld = cc.cc_day
            m = ld.dt.month
            d = ld.dt.day
            cc_from_color = cc.cc_from_color
            cc_to_color = cc.cc_to_color
            chg_type = cc.cc_type
            chg_descr = cc.cc_descr
            if (cmonth == -1):   #first month
                self.printPHP (3,"case {}:".format(m))
            elif (m != cmonth):
                self.printPHP (4,"break;")   # close previous month
                cmonth = cmonth + 1
                self.printPHP (3,"case {}:".format(cmonth)) # open new month
                if (ccolor != ColorType.GREEN):
                    self.printPHP (4,"$kleur= '{}';".format(ccolor))
                while (cmonth < m):     # if we're still not at month, close old and open new month
                    cmonth = cmonth + 1
                    self.printPHP (4,"break;")   # close previous month
                    self.printPHP (3,"case {}:".format(cmonth))
                    if (ccolor != ColorType.GREEN):
                        self.printPHP (4,"$kleur= '{}';".format(ccolor))

            cmonth = m

            self.printPHP(4,"# {}: {}".format(ld.dt.strftime("%a %e %b"), ld.descr))
            
            if (chg_descr):
                self.printPHP(4,"# {}".format(chg_descr))
            # Color change type
            if   (chg_type == ColorChangeType.UNTIL_INC):
                expr = '<='
                col = cc_from_color
            elif (chg_type == ColorChangeType.UNTIL_EXC):
                expr = '<'
                col = cc_from_color
            elif (chg_type == ColorChangeType.AFTER_INC):
                expr = '>='
                col = cc_to_color
            elif (chg_type == ColorChangeType.AFTER_EXC):
                expr = '>'
                col = cc_to_color
            else: # (chg_type == ColorChangeType.SINGLEDAY):
                expr = '=='
                col = cc_to_color

            self.printPHP (4,"if ($dag {} {}) {{ $kleur = '{}'; }}".format(expr,d,col))
            ccolor = cc_to_color
            cmonth = m
            cday = d
            
            # print (cc[0].strftime("%a %e %b, %H:%M"), end=" ")
            #print (cc_from_color, "after={}". format(chg_type))
        
        self.printPHP (4,"break;")   # close last month
        footer="""
   }
}

set_liturgical_color($litcol_light,$litcol_dark, $litcol_font);
?>
        """

        self.printPHP(0, footer)

        self.fd_php.close()

def main():

    year = date.today().year
    for y in range(year, year+2):
        cal = LiturgicalCalendar(y)
        cal.genPHPLiturgicalCalendar()
        cal.genTXTLiturgicalCalendar()
        cal.genXLSXLiturgicalCalendar()

if __name__ == "__main__":
    main()
