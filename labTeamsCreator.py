import os
import random
import easygui
import openpyxl

def read_excel(path) -> [str]:
    workbook = openpyxl.load_workbook(path)
    worksheet = workbook.worksheets[0]
    students = []
    for i in range(1,101):
        value = str(worksheet.cell(row = i, column = 1).value)
        if value != '' and value != 'None':
            students.append(value)
        else:
            break
    return students

def read_teams(path) -> ([[str]],[str]):
    workbook = openpyxl.load_workbook(path)
    worksheet = workbook.worksheets[0]
    teams = []
    allStudents = []
    for i in range(1,101):
        firstStudent = str(worksheet.cell(row = i, column = 1).value)
        if firstStudent == '' or firstStudent == 'None':
            break
        secondStudent = str(worksheet.cell(row=i, column=2).value)
        if secondStudent == '' or secondStudent == 'None':
            raise ValueError('Invalid team in retrieved lab ' + os.path.basename(path) + ': only found 1 student in the team.')
        thirdStudent = str(worksheet.cell(row=i, column=3).value)
        if thirdStudent == '' or thirdStudent == 'None':
            teams.append([firstStudent,secondStudent])
            allStudents.append(firstStudent)
            allStudents.append(secondStudent)
        else:
            teams.append([firstStudent,secondStudent,thirdStudent])
            allStudents.append(firstStudent)
            allStudents.append(secondStudent)
            allStudents.append(thirdStudent)
    return teams,allStudents

def write_teams(path,teams):
    if os.path.exists(path):
        os.remove(path)
    workbook = openpyxl.Workbook()
    worksheet = workbook.worksheets[0]
    for i in range(1,len(teams)+1):
        team = teams[i - 1]
        for j in range(1,len(team)+1):
            c = worksheet.cell(row = i,column = j).value = team[j-1]
    workbook.save(path)

def create_teams(number_of_labs = 8):
    folderPath = os.getcwd()
    number_of_labs = easygui.enterbox("How many labs?")
    if number_of_labs is None:
        return
    reuse = easygui.ynbox("Reuse labs already in folder?")
    try:
        if not number_of_labs.isnumeric():
            raise ValueError('Specified number of labs is not a number')
        number_of_labs = int(number_of_labs)
        xlsPath = os.path.join(folderPath, 'etudiants.xlsx')
        if not os.path.exists(xlsPath):
            raise ValueError('Could not find student file at expected location ' + xlsPath)
        etudiants = read_excel(xlsPath)
        pairings = {}
        for e in etudiants:
            pairings[e] = []
        labs = []

        for l in range(1,number_of_labs+1):
            teams = []
            if reuse:
                if os.path.exists(os.path.join(folderPath, str(l) + '.xlsx')):
                    oldTeams,oldStudents = read_teams(os.path.join(folderPath, str(l) + '.xlsx'))
                    if set(oldStudents) != set(etudiants):
                        raise ValueError('saved labs do not have the same students than current student file')
                    teams = oldTeams
                else:
                    reuse = False
            if len(teams) == 0:
                remainingStudents = [i for i in etudiants]
                random.shuffle(remainingStudents)
                attempts = 0
                while len(remainingStudents) != 0:
                    while len(remainingStudents) > 1:
                        firstStudent = remainingStudents.pop()
                        firstStudentPairings = pairings[firstStudent]
                        success = False
                        for i in range(len(remainingStudents)):
                            if remainingStudents[i] not in firstStudentPairings:
                                teams.append([firstStudent,remainingStudents.pop(i)])
                                success = True
                                break
                        if not success:
                            remainingStudents.append(firstStudent)
                            print('fail ' + str(attempts) + ' with remaining students ' + str(remainingStudents))
                            attempts += 1
                            if attempts % 10000 == 0:
                                raise ValueError('Found no solution after 10000 iterations. Aborting.')
                            elif attempts % 1000 == 0:
                                random.shuffle(teams)
                                remainingStudents += teams.pop()
                                remainingStudents += teams.pop()
                                remainingStudents += teams.pop()
                                attempts += 1
                                random.shuffle(remainingStudents)
                            elif attempts % 100 == 0:
                                random.shuffle(teams)
                                remainingStudents += teams.pop()
                                remainingStudents += teams.pop()
                            elif attempts % 10 == 0:
                                random.shuffle(teams)
                                remainingStudents += teams.pop()
                                random.shuffle(remainingStudents)
                            else:
                                random.shuffle(remainingStudents)

                    if len(remainingStudents) == 1:
                        lastStudent = remainingStudents.pop()
                        lastStudentPairings = pairings[lastStudent]
                        success = False
                        for team in teams:
                            if team[0] not in lastStudentPairings and team[1] not in lastStudentPairings:
                                team.append(lastStudent)
                                success = True
                                break
                        if not success:
                            remainingStudents.append(lastStudent)
                            print('fail ' + str(attempts) + ' with remaining students ' + str(remainingStudents))
                            random.shuffle(teams)
                            remainingStudents += teams.pop()
                            random.shuffle(remainingStudents)

            #success
            for team in teams:
                for student in team:
                    otherStudents = [s for s in team if s is not student]
                    pairings[student] += otherStudents
            labs.append(teams)

        labNb = 1
        for lab in labs:
            write_teams(os.path.join(folderPath,str(labNb)+'.xlsx'),lab)
            labNb += 1
    except Exception as e:
        easygui.msgbox(str(e),'error')
        raise e



if __name__ == '__main__':
    create_teams(10)