import xlsxwriter
import canvasapi
import pathlib
import sys
from canvasapi import Canvas
from cryptography.fernet import Fernet
from operator import itemgetter


def build_canvas():
    """
    Instantiate the Canvas object using the LMS URL and the
    APIKEY.txt file (you know to keep the key secret).
    """
    global canvas
    apiURL = "https://lms.au.af.edu"
    f = Fernet(key)  # This key is from the start of the main pgm
    apiKey = None
    while not apiKey:
        try:
            with open(apiKeyFile, 'rb') as file:  # We need to decrypt the text from the APIKEY.enc file
                encapiKEY = file.read()
                apiKey = f.decrypt(encapiKEY)  # We read this as bytes
                apiKey = apiKey.decode()  # So we need to decode back to a str
        except FileNotFoundError:
            print("File 'APIKEY.enc' not found!")
            create_apikeyfile()  # If the APIKEY.enc file isn't found, ask the user their key and save it
    file.close()
    canvas = Canvas(apiURL, apiKey)  # Create the canvas object
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
        accountID = input("\nType your Canvas account ID number and press enter, or \n"
                          "type 'help' for instructions to find your account ID:  ")
        print('\n')
        is_exit(accountID)
        if "help" in accountID:
            print("\nTo get your account ID, go to Canvas in your web browser and \n"
                  "navigate to a course. Click on 'People' and search for yourself.  \n"
                  "While hovering your mouse cursor over your name, look at the \n"
                  "bottom of the browser to see the hyperlink.  Your account ID is \n"
                  "the last number in the link.\n")
            continue
        break
    try:
        user = canvas.get_user(accountID)
        badUser = False
        return
    except canvasapi.exceptions.ResourceDoesNotExist:
        print("Could not find a user with ID", accountID)
        return
    except (canvasapi.exceptions.Unauthorized, canvasapi.exceptions.InvalidAccessToken) as e:
        print("Either the user ID", accountID, "is not authorized to use this API key, \n"
              "or there is an invalid API key in APIKEY.enc.\n")
        yesno = input("Would you like to change or update your API Key? (Yes/No):  ")
        fix_apikey(yesno)
        return
    except:
        print("Some sort of error occurred creating a user object for ", accountID)
        sys.exit(1)


def ask_course():
    """
    Ask which course they are looking to download from using the human readable
    name, then return with it.  Also, the user can type list for a list courses
    available to the user object.
    """
    global uniqueID
    print("Courses available to you: \n")
    for crse in courses:
        print(crse.name)
    while True:
        uniqueID = input("\nEnter the unique name for your course from above (i.e. 20-2 or 20-B)\n"
                         "or, type 'list' to display the list of courses available to you:  ")
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
            print("No course matching", uniqueID, "found.\n")
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
    global badAsgmt
    global wantedSubmissions
    print("Listing assignments for", course.name)
    for assignment in course.get_assignments():
        print(assignment)
    while True:
        print('\n')
        assignmentID = input("\nEnter the the assignment ID (in parentheses above) for the \n"
                             "grades you would like to retrieve, or type 'all' for all grades:  ")
        is_exit(assignmentID)
        break
    if "all" in str(assignmentID):
        for asgmt in course.get_assignments():
            assignmentID = str(asgmt.id)
            assignment = course.get_assignment(assignmentID)  # We're getting this for variable name uniqueness
            wantedSubmissions = course.get_multiple_submissions(student_ids='all', assignment_ids=str(asgmt.id),
                                                                include='rubric_assessment')
            badAsgmt = False
            get_rubric()
    else:
        try:
            assignment = course.get_assignment(assignmentID)  # We're getting this for variable name uniqueness
            wantedSubmissions = course.get_multiple_submissions(student_ids='all', assignment_ids=assignmentID,
                                                                include='rubric_assessment')
            badAsgmt = False
            get_rubric()
        except canvasapi.exceptions.ResourceDoesNotExist:
            print("\nCould not find the requested assignment", assignmentID)
            return
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
                break
            else:
                count += 1

    for rbrc in rbrcAsgmtMap:
        if assignmentID in str(rbrc):  # Assignment IDs are more unique than rubric IDs
            rubric = course.get_rubric(rbrc[0])
            break

    if rubric:
        canvas_rubrics()  # If we're good, let's do this thing
    else:
        print("No match for", assignment)
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
    print("Grabbing", assignment.name, "scores.")

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
            count = 0
            stuScores = [sub.user_id]  # Our list for the current student/submission, index[0] is the student ID
            if hasattr(sub, 'rubric_assessment'):  # Check if the student even has a submission/rubric assessment
                while count < len(sub.rubric_assessment):
                    for key in sub.rubric_assessment.keys():
                        if 'blank' in sub.rubric_assessment[key]['rating_id']:
                            stuScores.append('BLANK')
                            count += 1
                        else:
                            stuScores.append(sub.rubric_assessment[key]['points'])  # Append each rubric item score
                            count += 1
                stuScores.append(sub.grade)  # Append this student's overall score
                scoresAll.append(stuScores)  # Append this student's score list to the full list
    except:  # Catch unpublished assignments - or other errors *shrug*
        print("Error in processing", assignment.name, ". Is it published?  Skipping for now.")
        return
    scoresAll = sorted(scoresAll, key=itemgetter(0))  # Sort this list by student ID
    if len(scoresAll) == 0:
        print(assignment.name, " has no graded rubrics.  Skipping.")
        return

    flts = course.get_sections(include='students')  # We need a student/flight(section) list
    stdFltList = []
    for flt in flts:
        count = 0
        for s in flt.students:
            stdFltList.append([flt.students[count]['id'], flt.students[count]['sortable_name'], flt.name])
            count += 1
    xlsxOut = sorted(stdFltList, key=itemgetter(0))  # Sort this in the same manner as scoresALL

    scoresCount = 0
    xlsxCount = 0
    for item in xlsxOut:
        if scoresCount <= len(scoresAll) - 1:
            if item[0] == scoresAll[scoresCount][0]:  # Match student IDs before appending
                scoresAll[scoresCount].pop(0)  # Remove the redundant student ID
                xlsxOut[xlsxCount].extend(scoresAll[scoresCount])  # Place student's scores next to their name
                scoresCount += 1
                xlsxCount += 1
            else:
                xlsxOut[xlsxCount].append('NO SCORED RUBRIC FOUND ON CANVAS')
                xlsxCount += 1
        else:
            xlsxOut[xlsxCount].append('NO SCORED RUBRIC FOUND ON CANVAS')
            xlsxCount += 1

    xlsxOut.insert(0, rubricItems)  # Inserting rubric info at the top
    count = 0
    for item in rubricRatings:
        xlsxOut.insert(count, item)  # Inserting more rubric info at the top
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


def is_exit(exit):
    """
    Check if the  user types exit at any prompt.
    """
    if "exit" in exit:
        print("\nSee you later!")
        sys.exit(0)
    else:
        pass


def get_datadir() -> pathlib.Path:
    """
    Returns a parent directory path where persistent application data can be stored.
    This is the best location for the encryption key, all things considered.
    """
    home = pathlib.Path.home()

    if sys.platform == "win32":
        return home / "AppData/Roaming"
    elif sys.platform == "linux":
        return home / ".local/share"
    elif sys.platform == "darwin":
        return home / "Library/Application Support"


def load_key():
    """
    Returns the encryption/decryption key as a variable to use later.
    """
    try:
        return open(keyFile, "rb").read()
    except FileNotFoundError:
        print("Encryption key not found, generating a new one...")
        gen_key()


def gen_key():
    """
    Generate a new encryption key, and save it in the user's application
    data directory (userDir).
    """
    newKey = Fernet.generate_key()
    with open(keyFile, "wb") as file:
        file.write(newKey)
    return


def create_apikeyfile():
    """
    Write/overwrite the encrypted APIKEY.enc file with a new LMS API Key.
    """
    global key
    if not key:
        key = load_key()
    f = Fernet(key)
    while True:
        apiKeyIn = input("\nPlease enter the API key you generated on LMS or\n"
                         "type 'help' for instructions on generating a key:  ")
        is_exit(apiKeyIn)
        if "help" in apiKeyIn:
            print("\nTo generate an API Token/Key on LMS, login and click on\n"
                  "'Account'>'Settings'.  Under 'Approved Integrations' click\n"
                  "the '+New Access Token' button.  Fill the form, and click\n"
                  "the 'Generate Token' button.  A screen will pop up showing\n"
                  "your new token/key.  *You will not see that key again in its\n"
                  "entirety after clicking the 'X' in the top corner!* Write it\n"
                  "down in a **safe** place.  This token/key gives full LMS access\n"
                  "to any person that possesses it!  CanvasMyRubrics encrypts your\n"
                  "token when you enter it at the prompt.")
            continue
        break
    apiKeyIn = apiKeyIn.encode()  # We need to encode the API key into bytes for encryption
    encapiKeyIn = f.encrypt(apiKeyIn)
    with open(apiKeyFile, "wb") as file:
        file.write(encapiKeyIn)
    file.close()


def fix_apikey(yesno):
    if yesno.lower() == 'yes' or yesno.lower() == 'y':
        create_apikeyfile()
        build_canvas()
        return
    elif yesno.lower() == 'no' or yesno.lower() == 'n':
        return


print("""\nWelcome to CanvasMyRubrics.\nType exit at any prompt to exit the program.""")
userDir = get_datadir() / "CanvasMyRubrics"
try:
    userDir.mkdir(parents=True)
except FileExistsError:
    pass
keyFile = pathlib.Path(userDir / "APIKEY.key")
apiKeyFile = "APIKEY.enc"
key = None
while not key:
    key = load_key()

build_canvas()  # Open the Canvas API and create the canvas object

badUser = True  # This helps us continue to run build_user() when unexpected input is given
while badUser:
    build_user()  # Create a user object to get course info
key = None  # Clear the key to deter nefarious characters, should be done with it here anyway

courses = user.get_courses()  # We need the course list based on user for build_course() and ask_course()

badSearch = True  # This helps us continue to run ask_course() when unexpected input is given
while badSearch:
    build_course()  # Create a course object to get assignment info and rubrics

filename = uniqueID + ".xlsx"  # Create a filename based on the course name given by the user
workbook = xlsxwriter.Workbook(filename)  # Open a workbook - *this overwrites existing files with the same name*

badAsgmt = True
while badAsgmt:
    select_assignment()  # Get the user to select and assignment and call get_rubric and canvas_rubrics

workbook.close()  # Close the workbook

print("\nAll requested grades written to", filename, "in the current directory.")
sys.exit(0)
