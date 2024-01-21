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
        students = read_excel(xlsPath)
        pairings = {}
        studentsInTrios = []
        for e in students:
            pairings[e] = []
        labs = []
        reusedLabs = []

        while len(labs) + len(reusedLabs) < number_of_labs:
            teams = []
            if reuse:
                current_lab = len(reusedLabs) + 1
                if os.path.exists(os.path.join(folderPath, str(current_lab) + '.xlsx')):
                    oldTeams,oldStudents = read_teams(os.path.join(folderPath, str(current_lab) + '.xlsx'))
                    if set(oldStudents) != set(students):
                        raise ValueError('saved labs do not have the same students than current student file')
                    teams = oldTeams
                else:
                    reuse = False
            if len(teams) != 0:
                for team in teams:
                    if len(team) == 3:
                        studentsInTrios += team
                    for student in team:
                        otherStudents = [s for s in team if s is not student]
                        pairings[student] += otherStudents
                reusedLabs.append(teams)
                continue
            
            trioTeam = []
            preferredStudentsForTrio = []
            if len(students) - len(studentsInTrios) < 3:
                resetTrios = easygui.ynbox(
                    "Too many labs and not enough students, we need them to be in trios a second time. Allow it?")
                if not resetTrios:
                    raise ValueError(
                        'Too many labs and not enough students to have them participate in only one trio. Aborting.')
                else:
                    preferredStudentsForTrio = [s for s in students if s not in studentsInTrios]
                    studentsInTrios = []
            remainingStudents = [i for i in students]
            random.shuffle(remainingStudents)
            attempts = 0
            while len(remainingStudents) != 0:
                success = False
                firstStudent = None
                secondStudent = None
                if len(remainingStudents) % 2 != 0: # we build the trio first
                    if len(trioTeam) != 0:
                        raise ValueError('Programming error : trio team already exists')
                    if len(preferredStudentsForTrio) > 2:
                        raise ValueError('Programming error : preferredStudentForTrio should have 0, 1 or 2 elements')
                    if preferredStudentsForTrio:
                        firstStudent = preferredStudentsForTrio.pop()
                        remainingStudents.remove(firstStudent)
                    else:
                        for i in range(len(remainingStudents)):
                            if remainingStudents[i] not in studentsInTrios:
                                firstStudent = remainingStudents.pop(i)
                                break
                    if firstStudent is None:
                        raise ValueError('All students have already been in trios')
                    firstStudentPairings = pairings[firstStudent]
                    if preferredStudentsForTrio and preferredStudentsForTrio[0] not in firstStudentPairings:
                        secondStudent = preferredStudentsForTrio.pop()
                        remainingStudents.remove(secondStudent)
                    else:
                        for i in range(len(remainingStudents)):
                            if remainingStudents[i] not in studentsInTrios and \
                                    remainingStudents[i] not in firstStudentPairings:
                                secondStudent = remainingStudents.pop(i)
                                break
                    preferredStudentsForTrio = []
                    if secondStudent is not None:
                        thirdStudent = None
                        bothStudentsPairings = firstStudentPairings + pairings[secondStudent]
                        for i in range(len(remainingStudents)):
                            if remainingStudents[i] not in studentsInTrios and \
                                    remainingStudents[i] not in bothStudentsPairings:
                                thirdStudent = remainingStudents.pop(i)
                                break
                        if thirdStudent is not None:
                            trioTeam = [firstStudent, secondStudent, thirdStudent]
                            success = True
                else:
                    firstStudent = remainingStudents.pop()
                    firstStudentPairings = pairings[firstStudent]
                    for i in range(len(remainingStudents)):
                        if remainingStudents[i] not in firstStudentPairings:
                            teams.append([firstStudent,remainingStudents.pop(i)])
                            success = True
                            break
                if success:
                    attempts = attempts - (attempts % 5) #after a success, we always have 5 attempts before trying something else
                else:
                    attempts += 1
                    if firstStudent is not None:
                        remainingStudents.append(firstStudent)
                    if secondStudent is not None:
                        remainingStudents.append(secondStudent)
                    print('fail ' + str(attempts) + ' with remaining students ' + str(remainingStudents))
                    if attempts > 1e5:
                        raise ValueError('Found no solution after 100000 iterations. Aborting.')
                    elif attempts % 5 == 0:
                        #we neeed to pop teams
                        random.shuffle(teams)
                        milestones = [5,25,125,625,3125,15625,78125]
                        for milestone in milestones:
                            if attempts % milestone == 0:
                                if len(teams) > 0 :
                                    remainingStudents += teams.pop()
                                else:
                                    if trioTeam:
                                        remainingStudents += trioTeam
                                        trioTeam = []
                                    break
                    random.shuffle(remainingStudents)


            #success
            if trioTeam:
                studentsInTrios += trioTeam
                teams.append(trioTeam)
            for team in teams:
                for student in team:
                    otherStudents = [s for s in team if s is not student]
                    pairings[student] += otherStudents
            labs.append(teams)

        allLabs = reusedLabs + labs
        labNb = 1
        for lab in reusedLabs:
            labNb += 1
        for lab in labs:
            write_teams(os.path.join(folderPath,str(labNb)+'.xlsx'),lab)
            labNb += 1

    except Exception as e:
        easygui.msgbox(str(e),'error')
        raise e



if __name__ == '__main__':
    create_teams(10)