
import os
from xlrd import *
from easygui import *
import xlsxwriter
from datetime import *
import itertools




playLocations = ["1B", "2B", "3B", "HP", "LF", "CF", "RF"]

gamestate = {"A": ["Bases Empty Nobody Out", 1, 1],   "B": ["Runner on 1st, Nobody Out", 2, 4], "C": ["Runner on 2nd, Nobody Out", 3, 7], "D": ["Runner on 3rd, Nobody Out", 4, 10], "E": ["1st and 2nd, Nobody Out", 5, 13], "F": ["Runners at the Corners, Nobody Out", 6, 16],
             "G": ["2nd and 3rd, Nobody Out", 7, 19], "H": ["Bases Loaded, Nobody Out", 8, 22], "I": ["Bases Empty, One Out", 9, 2], "J": ["Runner on 1st, One Out", 10, 5], "K": ["Runner on 2nd, One Out", 11, 8], "L": ["Runner on 3rd, One Out", 12, 11], "M" : ["1st and 2nd, One Out", 13, 14], "N": ["Runners at the Corners, One Out", 14, 17],
             "O": ["2nd and 3rd, One Out", 15, 20], "P": ["Bases Loaded, One Out", 16, 23], "Q": ["Bases Empty, Two Outs", 17, 3], "R": ["Runner on 1st, Two Outs", 18, 6], "S": ["Runner on 2nd, Two Outs", 19, 9], "T": ["Runner on 3rd, Two Outs", 20, 12], "U": ["1st and 2nd, Two Outs", 21, 15], "V": ["Runners at the Corners, Two Outs", 22, 18], "W": ["2nd and 3rd, Two Outs", 23, 21], "X": ["Bases Loaded, Two Outs", 24, 24],}

reviewGrades = {"D": 10, "M": 7, "Z": 7, "A": 5, "Y": 3, "N": -5, "G": 1, "U": 1, "P": 1, "O": 1, "F": 1, "B": 1, "W": 1, "C": 1}



firtBasePlayTypes = ['Force Play', 'Tag Play', 'Touching a Base', 'Tag-Ups']
secondBasePlayTypes = ['Tag Play', 'Force Play', 'Slide Rule', 'Touching a Base', 'Tag-Ups']
thirdBasePlayTypes = ['Tag Play', 'Force Play', 'Touching a Base', 'Tag-Ups', "Slide rule"]
homePlatePlayTypes = ['Tag Play', 'Hit By Pitch', 'HP Collision Rule', 'Force Play', 'Time Play', 'Touching a Base']
LeftFieldPlayTypes = ['Potential Home Run', 'Fair / Foul', 'Catch / No Catch', 'Spectator Interference', 'Stadium Boundary Call', 'Ground-Rule Double']
CenterFieldPlayTypes = ['Potential Home Run', 'Catch / No Catch', 'Spectator Interference', 'Stadium Boundary Call', 'Ground-Rule Double']
RightFieldPlayTypes = ['Potential Home Run', 'Fair / Foul', 'Catch / No Catch', 'Spectator Interference', 'Stadium Boundary Call', 'Ground-Rule Double']



savePath = os.path.expanduser("~/Desktop/camevaldata.xlsx")
workbook = xlsxwriter.Workbook(savePath)
worksheet = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()
worksheet3 = workbook.add_worksheet()
worksheet4 = workbook.add_worksheet()
filename = fileopenbox(msg=None, title=None, default="*.xlsx")
book = open_workbook(filename)

sheet = book.sheet_by_index(2)

### FORMATS ####

worksheet2.set_column(2, 2, 26.7)
worksheet2.set_column(3, 45, 5)
worksheet2.set_column(0, 1, 5)
worksheet2.set_row(1, 42)

worksheet4.set_column(2, 2, 26.7)
worksheet4.set_column(3, 45, 5)
worksheet4.set_column(0, 1, 5)
worksheet4.set_row(1, 42)

bold = workbook.add_format({'bold': True})
bold.set_font_size(9)

boldwrap = workbook.add_format({'bold': True})
boldwrap.set_text_wrap()





def runStats():
    colz = 1
    rowzz = 0

    worksheet2.write(0, 2, "AVGS", bold)
    worksheet2.write(1,0,"Runners", bold)
    worksheet2.write(1, 1, "Bases", bold)
    for location in playLocations:

        for gamestateSingle in gamestate:
            camTypes = {"LH": 0, "MH\n(R)": 0, "HH": 0, "OHH\n(R)": 0, "L1B": 0, "L1B\n(In)": 0, "L1B\n(Out)": 0,
                        "M1B": 0,
                        "H1B": 0, "L3B": 0,
                        "L3B\n(In)": 0, "L3B\n(Out)": 0, "M3B": 0, "H3B": 0, "LF\nSplash": 0, "LF\nPole": 0,
                        "LF\nAlley": 0,
                        "LF": 0, "CF": 0, "TCF": 0, "RF": 0, "RF\nAlley": 0,
                        "RF\nPole": 0, "RF\nSplash": 0, "Hand\nHeld": 0, "LH\n(Mo)": 0, "L1B (Mo)": 0,
                        "L1B (In) (Mo)": 0,
                        "L1B (Out) (Mo)": 0, "M1B SSMO": 0, "H1B SSMO": 0,
                        "L3B\n(Mo)": 0, "L3B (In) (Mo)": 0, "L3B (Out) (Mo)": 0, "M3B SSMO": 0, "H3B SSMO": 0,
                        "LF SSMO": 0,
                        "CF SSMO": 0, "TCF SSMO": 0, "RF SSMO": 0, "Dugout" : 0, "Trop\nRing\nRobo": 0, "Beauty": 0}

            camTotals = {"LH": 0, "MH\n(R)": 0, "HH": 0, "OHH\n(R)": 0, "L1B": 0, "L1B\n(In)": 0, "L1B\n(Out)": 0,
                         "M1B": 0,
                         "H1B": 0, "L3B": 0,
                         "L3B\n(In)": 0, "L3B\n(Out)": 0, "M3B": 0, "H3B": 0, "LF\nSplash": 0, "LF\nPole": 0,
                         "LF\nAlley": 0,
                         "LF": 0, "CF": 0, "TCF": 0, "RF": 0, "RF\nAlley": 0,
                         "RF\nPole": 0, "RF\nSplash": 0, "Hand\nHeld": 0, "LH\n(Mo)": 0, "L1B (Mo)": 0,
                         "L1B (In) (Mo)": 0,
                         "L1B (Out) (Mo)": 0, "M1B SSMO": 0, "H1B SSMO": 0,
                         "L3B\n(Mo)": 0, "L3B (In) (Mo)": 0, "L3B (Out) (Mo)": 0, "M3B SSMO": 0, "H3B SSMO": 0,
                         "LF SSMO": 0,
                         "CF SSMO": 0, "TCF SSMO": 0, "RF SSMO": 0, "Dugout" : 0, "Trop\nRing\nRobo": 0, "Beauty": 0}
            rowzz +=1


            for row in range(sheet.nrows):
                col = 26
                camGrade = sheet.cell_value(row, 5)

                if camGrade == gamestateSingle:
                    # print("this row has " + gamestateSingle, row)



                    if sheet.cell_value(row, 17) == location:
                        # print("this is at " + location)


                        while col < 154:
                            camLetter =  sheet.cell_value(row, col)
                            if camLetter is float:
                                camLetter.upper()
                            camName = sheet.cell_value(1, col)
                            if camLetter != "":
                                print(camLetter, col, camName )
                                if camName in camTypes:
                                    if camLetter in reviewGrades:
                                        camValue = reviewGrades[camLetter]
                                        # print("success!")
                                        # print(camValue, type(camValue))
                                        # print(camTypes[camName], camName, type(camTypes[camName]))
                                        camTypes[camName] = camTypes[camName] + camValue
                                        camTotals[camName] = camTotals[camName] + 1



                                    else:
                                        print("this is an odd camLetter ", camLetter, row)

                            col += 1

            # print("game status - " + gamestateSingle + " - " + gamestate[gamestateSingle] + " @ " + location, camTypes)
            # print("cam totals ", camTotals)

            print("\n\n")
            print("game status - " + gamestateSingle + " - " + gamestate[
                gamestateSingle][0] + " play location @ " + location)

            worksheet.write(0, colz, gamestateSingle + "-" + gamestate[gamestateSingle][0] + " @ "+ location)
            worksheet.write(2, colz, "CAMERA")
            worksheet.write(2, colz + 2, "TOTAL CAM")
            worksheet.write(2, colz + 3, "TOTAL SCORE")
            worksheet.write(2, colz + 4, "AVG SCORE")


            worksheet2.write(rowzz + 1, 0, gamestate[gamestateSingle][1])
            worksheet2.write(rowzz+1, 1, gamestate[gamestateSingle][2])
            worksheet2.write(rowzz+1, 2, gamestateSingle + "-" + gamestate[gamestateSingle][0] + " @ " + location, bold)
            rowz = 2
            colzz = 3
            for camera in camTypes:

                # print(camera)
                totalScore = camTypes[camera]
                totalCams = camTotals[camera]
                if totalCams != 0:
                    avgScore = totalScore / totalCams
                    avgScore = round(avgScore, 2)
                else:
                    # print(camera + " has total cams equal to 0")
                    avgScore = 0



                print(camera, "total score: ", totalScore, "totalCams: ", totalCams, "avg Score ", avgScore)


                worksheet.write(rowz + 1, colz, camera)
                worksheet.write(rowz + 1, colz + 2, totalCams)
                worksheet.write(rowz + 1, colz + 3, totalScore)
                worksheet.write(rowz+ 1, colz + 4, avgScore)


                worksheet2.write(1, colzz, camera, boldwrap)
                worksheet2.write(rowzz+1, colzz, avgScore)

                colzz +=1

                rowz += 1
            colz += 6
        rowzz +=2



def runStatsPerPlayType():
    colz = 1
    rowzz = 0

    worksheet4.write(0,2,"AVGS", bold)
    worksheet4.write(1, 0, "Runners", bold)
    worksheet4.write(1, 1, "Bases", bold)

    for location in playLocations:


        if location == '1B':
            plays = firtBasePlayTypes
        if location == '2B':
            plays = secondBasePlayTypes
        if location == '3B':
            plays = thirdBasePlayTypes

        if location == 'HP':
            plays = homePlatePlayTypes
        if location == 'LF':
            plays == LeftFieldPlayTypes
        if location == 'CF':
            plays = CenterFieldPlayTypes
        if location == 'RF':
            plays = RightFieldPlayTypes

        for play in plays:

            for gamestateSingle in gamestate:
                camTypes = {"LH": 0, "MH\n(R)": 0, "HH": 0, "OHH\n(R)": 0, "L1B": 0, "L1B\n(In)": 0, "L1B\n(Out)": 0,
                            "M1B": 0,
                            "H1B": 0, "L3B": 0,
                            "L3B\n(In)": 0, "L3B\n(Out)": 0, "M3B": 0, "H3B": 0, "LF\nSplash": 0, "LF\nPole": 0,
                            "LF\nAlley": 0,
                            "LF": 0, "CF": 0, "TCF": 0, "RF": 0, "RF\nAlley": 0,
                            "RF\nPole": 0, "RF\nSplash": 0, "Hand\nHeld": 0, "LH\n(Mo)": 0, "L1B (Mo)": 0,
                            "L1B (In) (Mo)": 0,
                            "L1B (Out) (Mo)": 0, "M1B SSMO": 0, "H1B SSMO": 0,
                            "L3B\n(Mo)": 0, "L3B (In) (Mo)": 0, "L3B (Out) (Mo)": 0, "M3B SSMO": 0, "H3B SSMO": 0,
                            "LF SSMO": 0,
                            "CF SSMO": 0, "TCF SSMO": 0, "RF SSMO": 0, "Dugout" : 0, "Trop\nRing\nRobo": 0, "Beauty": 0}

                camTotals = {"LH": 0, "MH\n(R)": 0, "HH": 0, "OHH\n(R)": 0, "L1B": 0, "L1B\n(In)": 0, "L1B\n(Out)": 0,
                             "M1B": 0,
                             "H1B": 0, "L3B": 0,
                             "L3B\n(In)": 0, "L3B\n(Out)": 0, "M3B": 0, "H3B": 0, "LF\nSplash": 0, "LF\nPole": 0,
                             "LF\nAlley": 0,
                             "LF": 0, "CF": 0, "TCF": 0, "RF": 0, "RF\nAlley": 0,
                             "RF\nPole": 0, "RF\nSplash": 0, "Hand\nHeld": 0, "LH\n(Mo)": 0, "L1B (Mo)": 0,
                             "L1B (In) (Mo)": 0,
                             "L1B (Out) (Mo)": 0, "M1B SSMO": 0, "H1B SSMO": 0,
                             "L3B\n(Mo)": 0, "L3B (In) (Mo)": 0, "L3B (Out) (Mo)": 0, "M3B SSMO": 0, "H3B SSMO": 0,
                             "LF SSMO": 0,
                             "CF SSMO": 0, "TCF SSMO": 0, "RF SSMO": 0, "Dugout" : 0, "Trop\nRing\nRobo": 0, "Beauty": 0}
                rowzz += 1

                for row in range(sheet.nrows):
                    col = 26
                    camGrade = sheet.cell_value(row, 5)

                    if camGrade == gamestateSingle:
                        # print("this row has " + gamestateSingle, row)



                        if sheet.cell_value(row, 17) == location:
                            # print("this is at " + location)

                            if sheet.cell_value(row, 16) == play:

                                while col < 154:
                                    camLetter = sheet.cell_value(row, col)
                                    if camLetter is float:
                                        camLetter.upper()
                                    camName = sheet.cell_value(1, col)
                                    if camLetter != "":
                                        # print(camLetter, col, camName )
                                        if camName in camTypes:
                                            if camLetter in reviewGrades:
                                                camValue = reviewGrades[camLetter]
                                                # print("success!")
                                                # print(camValue, type(camValue))
                                                # print(camTypes[camName], camName, type(camTypes[camName]))
                                                camTypes[camName] = camTypes[camName] + camValue
                                                camTotals[camName] = camTotals[camName] + 1



                                            else:
                                                print("this is an odd camLetter ", camLetter, row)

                                    col += 1

                # print("game status - " + gamestateSingle + " - " + gamestate[gamestateSingle] + " @ " + location, camTypes)
                # print("cam totals ", camTotals)

                print("\n\n")
                print("game status - " + gamestateSingle + " - " + gamestate[
                    gamestateSingle][0] + " play location @ " + location)
                worksheet3.write(0, colz, gamestateSingle + "-" + gamestate[gamestateSingle][0] + " @ " + location + " " + play)
                worksheet3.write(2, colz, "CAMERA")
                worksheet3.write(2, colz + 2, "TOTAL CAM")
                worksheet3.write(2, colz + 3, "TOTAL SCORE")
                worksheet3.write(2, colz + 4, "AVG SCORE")

                worksheet4.write(rowzz + 1, 2, gamestateSingle + "-" + gamestate[gamestateSingle][0] + " @ " + location + " " + play, bold)
                rowz = 2
                colzz = 3
                for camera in camTypes:

                    # print(camera)
                    totalScore = camTypes[camera]
                    totalCams = camTotals[camera]
                    if totalCams != 0:
                        avgScore = totalScore / totalCams
                        avgScore = round(avgScore, 2)
                    else:
                        # print(camera + " has total cams equal to 0")
                        avgScore = 0

                    print(camera, "total score: ", totalScore, "totalCams: ", totalCams, "avg Score ", avgScore)

                    worksheet3.write(rowz + 1, colz, camera)
                    worksheet3.write(rowz + 1, colz + 2, totalCams)
                    worksheet3.write(rowz + 1, colz + 3, totalScore)
                    worksheet3.write(rowz + 1, colz + 4, avgScore)

                    worksheet4.write(rowzz + 1, 0, gamestate[gamestateSingle][1])
                    worksheet4.write(rowzz + 1, 1, gamestate[gamestateSingle][2])
                    worksheet4.write(1, colzz, camera, boldwrap)
                    worksheet4.write(rowzz + 1, colzz, avgScore)

                    colzz += 1

                    rowz += 1
                colz += 6
            rowzz +=2

runStats()
runStatsPerPlayType()

workbook.close()