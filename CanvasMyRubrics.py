import xlsxwriter
from canvasapi import Canvas
from sys import exit
from operator import itemgetter


def build_canvas():
    """
    Instantiate the Canvas object using the LMS URL and the
    APIKEY.txt file (you know to keep the key secret).
    """
    global canvas
    apiURL = "https://lms.au.af.edu"
    try:
        with open('APIKEY.txt', 'r') as file:
            apiKEY = file.read()
    except FileNotFoundError:
        print("File 'APIKEY.txt' not found! Make sure you place it in the same \n"
              "folder as this executable.  Exiting.")
        exit(1)
    file.close()
    canvas = Canvas(apiURL, apiKEY)
    return


def build_user():
    """
    Instantiate a user object using the user's LMS ID...no easy way to get a list of
    IDs. So they need to retrieve that before using this program. A help option
    is provided.
    """
    global badUser
    global user
    while True:
        accountID = input("Type your Canvas account ID number and press enter, or \n"
                          "type 'help' for instructions to find your account ID:  ")
        print('\n')
        is_exit(accountID)
        if "help" in accountID:
            print("\nTo get your account ID, go to Canvas in your web browser and \n"
                  "navigate to a course. Click on 'People' and search for yourself.  \n"
                  "While hovering your mouse cursor over your name, look at the \n"
                  "bottom of the browser to see a hyperlink.  Your account ID is the \n"
                  "last number in the link.\n")
            continue
        break
    try:
        user = canvas.get_user(accountID)
        badUser = False
    except:
        print("Are you sure " + str(accountID) + " is the correct ID number?")
    return


def ask_course():
    """
    Ask which course they are looking to download from using the human readable
    name, then return with it.  Also, the user can type list for a list courses
    available to the user object.
    """
    global uniqueID
    while True:
        uniqueID = input("Enter the unique name for your course (i.e. 20-2 or 20-B), \n"
                         "or type 'list' for a list of courses available to you:  ")
        print('\n')
        is_exit(uniqueID)
        if "list" in uniqueID:
            for crse in courses:
                print(crse.name)
                continue
        else:
            return uniqueID


def build_course():
    """
    Instantiate a course object using the ask_course() function to search the courses
    available to the user object.
    """
    global badSearch
    global course
    ask_course()
    count = 1
    coursesLen = len(courses._elements)
    for crse in courses:
        if uniqueID in crse.name:
            badSearch = False
            courseID = crse.id
        elif count == coursesLen:
            print("No course matching", uniqueID, "found.")
            return
        else:
            count += 1
    course = canvas.get_course(courseID)
    return


def select_assignment():
    """
    List all of the assignments the course object has access to. Then, using the course
    object, create an the assignment object requested by the user.  Finally, call
    get_rubric() do iterate, download, and write the requested rubrics to an xlsx file.
    """
    global assignment
    global assignmentID
    global wantedSubmissions
    print("Listing assignments for", course.name)
    for assignment in course.get_assignments():
        print(assignment)
    while True:
        print('\n')
        assignmentID = input("Enter the the assignment ID (in parentheses above) for the \n"
                             "grades you would like to retrieve, or type 'all' for all grades:  ")
        is_exit(assignmentID)
        break
    if "all" in str(assignmentID):
        for asgmt in course.get_assignments():
            wantedSubmissions = course.get_multiple_submissions(student_ids='all', assignment_ids=str(asgmt.id),
                                                                include='rubric_assessment')
            assignmentID = str(asgmt.id)
            assignment = course.get_assignment(assignmentID)  # We're getting this for variable name uniqueness
            get_rubric()
    else:
        wantedSubmissions = course.get_multiple_submissions(student_ids='all', assignment_ids=assignmentID,
                                                            include='rubric_assessment')
        assignment = course.get_assignment(assignmentID)  # We're getting this for variable name uniqueness
        get_rubric()
    return


def get_rubric():
    """
    Get important rubric (not assignment or submission) information and map rubric
    IDs to assignment IDs.  This allows us to match the rubric scoring and item info
    to an assignment (and later a submission).
    """
    global rubrics
    global rubric
    rubric = None
    rubrics = course.get_rubrics()
    rubricsList = []
    for rbrc in rubrics:
        rubricsList.append(rbrc.title)

    rbrcAsgmtMap = []
    for asgmt in course.get_assignments():
        count = 0
        while count < len(rubricsList):  # Below...current naming conventions are unambiguous in the first 10 chars
            if rubrics[count].title[0:10] in asgmt.name[0:10]:  # They also won't match if you go too far
                rbrcAsgmtMap.append([rubrics[count].id, asgmt.id])
            count += 1

    for rbrc in rbrcAsgmtMap:
        if assignmentID in str(rbrc):  # Assignment IDs are more unique than rubric IDs
            rubric = course.get_rubric(rbrc[0])

    if rubric:
        canvas_rubrics()  # If we're good, let's do this thing
    else:
        print("No match for " + str(assignment))
        return


def canvas_rubrics():
    """
    Build list of lists containing rubric info, assignment info, student/flight info, and
    submission grades for the selected assignment.  Uses that list to write to the
    xlsx file.
    """
    global worksheet
    global studentList
    global submission
    global scoresAll
    global xlsxOut
    print("Grabbing" + ' ' + assignment.name + " scores.")

    rubricItems = ['Student ID', 'Student Name', 'Flight']  # Establish column headers before rubric items
    rubricRatings = [[], [], []]  # List of lists containing rubric rating options/points for each item

    count = 0
    for item in rubric.data:  # Go through the rubrics to finish column headers and rating points
        rubricItemDesc = rubric.data[count]['description']
        rubricItemPoints = rubric.data[count]['points']
        rubricItems.append(rubricItemDesc + ' ' + str(rubricItemPoints))
        for rbrc in range(3):
            ratingPoints = rubric.data[count]['ratings'][rbrc]['points']
            rubricRatings[rbrc].append(ratingPoints)
        count += 1
    rubricItems.append('Highest possible: ' + str(rubric.points_possible))

    rubricRatingDesc = ['Exceeds', 'Meets', 'Does Not Meet']  # Place rating options in front of points
    for rating in range(3):
        rubricRatings[rating].insert(0, '')
        rubricRatings[rating].insert(0, rubricRatingDesc[rating])
        rubricRatings[rating].insert(0, '')

    scoresAll = []  # Our full list of scores for each submission
    try:  # We will get an exception if the assignment isn't published
        for sub in wantedSubmissions:
            if hasattr(sub, 'rubric_assessment'):  # Check if the student even has a submission/rubric assessment
                count = 0
                stuScores = [sub.user_id]  # Our list for the current student/submission, first index is the student ID
                while count < len(sub.rubric_assessment):
                    for key in sub.rubric_assessment.keys():
                        stuScores.append(sub.rubric_assessment[key]['points'])  # Append individual rubric item scores
                        count += 1
                stuScores.append(sub.grade)  # Append this student's overall score
                scoresAll.append(stuScores)  # Append this student's score list to the full list
    except:
        print(assignment.name + " is probably not published...Skipping.")  # Catch unpublished assignments
        return
    scoresAll = sorted(scoresAll, key=itemgetter(0))  # Sort this list by student ID
    for sub in scoresAll:  # Remove the student ID since we're mapping it to a section below
        sub.pop(0)

    flts = course.get_sections(include='students')  # We need a student/flight(section) list
    stdFltList = []
    for flt in flts:
        count = 0
        for s in flt.students:
            stdFltList.append([flt.students[count]['id'], flt.students[count]['sortable_name'], flt.name])
            count += 1
    xlsxOut = sorted(stdFltList, key=itemgetter(0))  # Sort this in the same manner as scoresALL

    count = 0
    for item in scoresAll:
        xlsxOut[count].extend(scoresAll[count])  # Place student's scores next to their name
        count += 1
    xlsxOut.insert(0, rubricItems)

    count = 0
    for item in rubricRatings:
        xlsxOut.insert(count, item)  # Inserting rubric info at the top
        if count == 1:
            xlsxOut[count].append('Minimum Passing Score')
            count += 1
        elif count == 2:
            xlsxOut[count].append(str(rubric.points_possible * .7))
        else:
            count += 1

    if len(rubric.title) >= 25:  # Worksheets have a name length limit of 31 chars and we need unique names
        worksheetName = str(assignment.name[0:24]) + str(assignmentID)
    else:
        worksheetName = str(assignment.name) + str(assignmentID)
    worksheet = workbook.add_worksheet(worksheetName)  # Make a new worksheet
    row_writer(xlsxOut)  # Write it out to the file


def row_writer(data):
    """
    This builds and writes simple rows in the xlsx from python list of lists.
    """
    row = 0
    while row < len(data):
        col = 0
        for item in data[row]:
            worksheet.write(row, col, item)
            col += 1
        row += 1
    return


def is_exit(x):
    """
    Check if the  user types exit at any prompt.
    """
    if "exit" in x:
        print("See you later!")
        exit(0)
    else:
        pass


print("""\nWelcome to this rubric downloader.\nType exit at any prompt to exit the program.\n""")
build_canvas()  # Open the Canvas API and create the canvas object
badUser = True  # This helps us continue to run build_user() when unexpected input is given
while badUser:
    build_user()  # Create a user object to get course info

courses = user.get_courses()  # We need the course list based on user for build_course() and ask_course()

badSearch = True  # This helps us continue to run ask_course() when unexpected input is given
while badSearch:
    build_course()  # Create a course object to get assignment info and rubrics

filename = uniqueID + ".xlsx"  # Create a filename based on the course name given by the user
workbook = xlsxwriter.Workbook(filename)  # Open a workbook - *this overwrites existing files with the same name*

select_assignment()  # Get the user to select and assignment and call get_rubric and canvas_rubrics

workbook.close()  # Close the workbook
exit(0)
