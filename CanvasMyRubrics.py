import xlsxwriter
from canvasapi import Canvas
from sys import exit


def is_exit(x):
    """
    Check if the  user types exit at any prompt.
    """
    if "exit" in x:
        print('See you later!')
        exit(0)
    else:
        pass


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
    except:
        print("""File 'APIKEY.txt' not found! Make sure you place it in the same 
folder as this executable.  Press Enter or Return to exit.""")
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
    global user
    while True:
        accountID = input("""Type your Canvas account ID number and press enter,
or type "help" for instructions to find your account ID:  """)
        print('\n')
        is_exit(accountID)
        if "help" in accountID:
            print("""\nTo get your account ID, go to Canvas in your web browser and
navigate to a course. Click on "People" and search for yourself.  While hovering 
your mouse cursor over your name, look at the bottom of the browser to see
a hyperlink.  Your account ID is the last number in the link.\n""")
            continue
        break
    user = canvas.get_user(accountID)
    return


"""
Instantiate a course object using the ask_course() function to search the courses
available to the user object.
"""


def build_course():
    global badSearch
    global course
    ask_course()
    count = 1
    pCoursesLen = len(pCourses.__dict__)
    while count < pCoursesLen:
        for c in pCourses:
            if uniqueID in c.name:
                badSearch = False
                course = canvas.get_course(c.id)
                return
            else:
                count += 1
                continue
    print('No course matching', uniqueID, 'found.')
    return


def ask_course():
    """
    Ask which course they are looking to download from using the human readable
    name, then grab the unique Canvas ID and return with it.  Also, the user can
    type list for a list courses available to the user object.
    """
    global uniqueID
    while True:
        uniqueID = input("""Enter the unique name for your course (i.e. 20-2 or 20-B),
or type "list" for a list of courses available to you:  """)
        print('\n')
        is_exit(uniqueID)
        if "list" in uniqueID:
            for c in pCourses:
                print(c.name)
                continue
        return uniqueID


def build_assignment():
    """
    List all of the assignments the course object has access to. Then, using the course
    object, create an the assignment object requested by the user.  Finally, call canvas_rubrics()
    do iterate, download, and write the requested rubrics to an xlsx file.
    """
    global assignment
    print('Listing assignments for', course.name)
    for a in course.get_assignments():
        print(a)
    while True:
        print('\n')
        assignmentID = input("""Enter the the assignment ID (in parentheses above) for rubric 
would you like to retrieve, or type "all" for all rubrics:  """)
        is_exit(assignmentID)
        break
    if "all" in assignmentID:
        for a in course.get_assignments():
            assignment = course.get_assignment(a.id)
            canvas_rubrics()
    else:
        assignment = course.get_assignment(assignmentID)
        canvas_rubrics()


def write_headers():
    """
    Here we write the headers to the xlsx file using static placement for the student name and flight
    name.  Because Canvas makes unique codes for everything and the rubrics are of varying lenghts,
    we'll make friendly "item" names for each rubric item.
    """
    headers = list(submission.rubric_assessment.keys())
    count = 0
    for h in headers:
        headers[count] = "Item" + str(count + 1)
        count += 1
    headers.insert(0, "Flight")
    headers.insert(0, "Student Name")
    row_writer(0, 0, headers)
    return


def row_writer(row, col, data):
    """
    This builds and writes simple rows in the xlsx from simple python lists.
    """
    count = 0
    for i in data:
        worksheet.write(row, col, data[count])
        col += 1
        count += 1
    return


def flight_mapper():
    """
    Map the student name to their flight name (section in Canvas speak), and fill the first two columns
    using the studentList variable.  We run this once per worksheet because it's time intensive.
    """
    print('Mapping students to flights for the first two columns.')
    flightMap = {}
    flightList = course.get_sections(include="students")
    for s in studentList:
        for f in flightList:
            n = len(f.students)
            for i in range(0, n):
                if str(s.id) in str(f.students[i]['id']):
                    flightMap.update({s.display_name: f.name})
    row = 1
    col = 0
    for name, flight in flightMap.items():
        worksheet.write(row, col, name)
        worksheet.write(row, col + 1, flight)
        row += 1
    return


def canvas_rubrics():
    """
    This is where all of the rubric data is discovered, iterated, and writen to the xlsx file.
    """
    global worksheet
    global studentList
    global submission
    studentList = assignment.get_gradeable_students()

    if len(assignment.name) >= 25:  # Worksheets have a name length limit of 31 chars and we need unque names
        assignName = str(assignment.name[0:24]) + str(assignment.id)
    else:
        assignName = str(assignment.name) + str(assignment.id)
    worksheet = workbook.add_worksheet(assignName)

    row = 1
    flight_mapper()
    headerWritten = False
    for s in studentList:
        submission = assignment.get_submission(s.id, include="rubric_assessment")
        if hasattr(submission,
                   "rubric_assessment"):  # We need to make sure there is a submitted rubric before attempting anything
            if headerWritten is False:
                keys = list(submission.rubric_assessment.keys())
                write_headers()
                headerWritten = True
            print('Grabbing rubric for', s, '...')
            scoreRow = []
            count = 0
            for k in keys:
                key = keys[count]
                scoreRow.append(submission.rubric_assessment[key]['points'])
                count += 1
            row_writer(row, 2, scoreRow)
            row += 1
        else:
            print('No rubric found for', s, '...')
            continue
    print('All submitted rubrics for', assignment.name, 'saved to', filename, '.')
    return


print("""\nWelcome to this rubric downloader.\nType exit at any prompt to exit the program.\n""")
build_canvas()  # Open the Canvas API and create the canvas object
build_user()  # Create a user object to get course info

pCourses = user.get_courses()  # We need the course list based on user for build_course() and ask_course()

badSearch = True  # This helps us continue to run ask_course() when unexpected input is given
while badSearch:
    build_course()  # Create a course object to get assignment info and rubrics
filename = uniqueID + ".xlsx"  # Create a filename based on the course name given by the user
workbook = xlsxwriter.Workbook(filename)  # Open a workbook - *this overwrites existing files with the same name*

build_assignment()  # Create an assignment object and call canvas_rubrics() as many times as necessary

workbook.close()
exit(0)
