### This file contains a number of utility functions that can be used in Tikit TMS scripts. ####

import clr
import datetime

clr.AddReference('mscorlib')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Windows.Forms')
import System

from datetime import date, timedelta
from System import DateTime
from System.Diagnostics import Process
from System.Globalization import DateTimeStyles
from System.Collections.Generic import Dictionary
from System.Windows import Controls, Forms, LogicalTreeHelper
from System.Windows import Data, UIElement, Visibility, Window
from System.Windows.Controls import Button, Canvas, GridView, GridViewColumn, ListView, Orientation
from System.Windows.Data import Binding
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import Brush, Brushes
##### constants ########################################################

### SQL tools ####################################################

def runSQL(codeToRun, showError = False, errorMsgText = "", errorMsgTitle = "", apostropheHandle = 0):
  # This function is written to handle and check inputted SQL code, and will return the result of the SQL code.
  # It first checks the length and wrapping of the code, then attempts to execute the SQL, it has an option apostrophe handler.
  # codeToRun     = Full SQL of code to run. No need to wrap in '[SQL: code_Here]' as we can do that here
  # showError     = True / False. Indicates whether or not to display message upon error
  # errorMsgText  = Text to display in the body of the message box upon error (note: actual SQL will automatically be included, so no need to re-supply that)
  # errorMsgTitle = Text to display in the title bar of the message box upon error

  if len(codeToRun) < 10:
    MessageBox.Show("The supplied 'codeToRun' doesn't appear long enough, please check and update this code if necessary.\nPassed SQL: " + str(codeToRun), "ERROR: runSQL...")
    return
  
  if codeToRun[:5] == "[SQL:":
    fCodeToRun = codeToRun
  else:
    fCodeToRun = "[SQL: " + codeToRun + "]"
  
  try:
    tmpValue = _tikitResolver.Resolve(fCodeToRun)
    if apostropheHandle == 1:
      tmpValue = tmpValue.replace("'", "''")
    returnVal = str(tmpValue)
    returnVal1 = 'N/A' if returnVal == None else returnVal
  except:
    if showError == True:
      MessageBox.Show("{0}\n\nSQL used:\n{1}".format(errorMsgText, codeToRun), errorMsgTitle)
    returnVal = ''
    returnVal1 = "!Error"
  
  if showError:
    msgBody = "runSQL(...):\n  CodeToRun: {0}\n  ShowError: {1}\n  ErrorMsgText: '{2}'\n  ErrorMsgTitle: '{3}'\n  > Result: {4}".format(fCodeToRun, showError, errorMsgText, errorMsgTitle, returnVal1)
    MessageBox.Show(msgBody)
  
  return returnVal1

# def runSQL(codeToRun, 
#            showError=False, 
#            errorMsgText="", 
#            errorMsgTitle="", 
#            apostropheHandle=0, 
#            runInThread=False):
#     """
#     Runs SQL by calling _tikitResolver.Resolve in either the current
#     thread or a separate .NET Thread (if runInThread=True).

#     :param codeToRun: SQL code to run.
#     :param showError: If True, show a message box on error.
#     :param errorMsgText: Text for the error message box body.
#     :param errorMsgTitle: Text for the error message box title.
#     :param apostropheHandle: If 1, replace single quotes in the resolved result.
#     :param runInThread: If True, runs the resolver on a separate .NET Thread (blocks until done).
#     """
#     # Quick length check
#     if len(codeToRun) < 10:
#         MessageBox.Show(
#             "The supplied 'codeToRun' doesn't appear long enough, "
#             "please check and update this code if necessary.\nPassed SQL: " + str(codeToRun),
#             "ERROR: runSQL..."
#         )
#         return

#     # Ensure we wrap the code if it is not already wrapped in [SQL: ...]
#     if codeToRun.startswith("[SQL:"):
#         fCodeToRun = codeToRun
#     else:
#         fCodeToRun = "[SQL: " + codeToRun + "]"

#     def resolve_sql():
#         """Encapsulates the call to _tikitResolver.Resolve(...) in a try/except block."""
#         try:
#             tmpValue = _tikitResolver.Resolve(fCodeToRun)
#             if apostropheHandle == 1 and tmpValue is not None:
#                 tmpValue = tmpValue.replace("'", "''")

#             # Convert None -> 'N/A', else str(tmpValue)
#             resolved = "N/A" if tmpValue is None else str(tmpValue)
#             return resolved
#         except Exception:
#             if showError:
#                 MessageBox.Show(
#                     "{0}\n\nSQL used:\n{1}".format(errorMsgText, codeToRun), 
#                     errorMsgTitle
#                 )
#             return "!Error"

#     if runInThread:
#         result_container = [None]  # store the result from worker

#         def worker():
#             result_container[0] = resolve_sql()

#         # Create a .NET thread using ThreadStart
#         thread = Thread(ThreadStart(worker))
#         thread.Start()
#         thread.Join()  # block until the thread finishes

#         returnVal1 = result_container[0]
#     else:
#         # Run synchronously
#         returnVal1 = resolve_sql()

#     # Debug info
#     if showError:
#         msgBody=(
#             "runSQL(...):\n"
#             "  CodeToRun: {0}\n"
#             "  ShowError: {1}\n"
#             "  ErrorMsgText: '{2}'\n"
#             "  ErrorMsgTitle: '{3}'\n"
#             "  > Result: {4}"
#         ).format(
#             fCodeToRun, showError, errorMsgText, errorMsgTitle, returnVal1
#         )
#         MessageBox.Show(msgBody)
#     return returnVal1

### Approval/access checks #############################################

def isUserAnApprovalUser(userToCheck):
  # This is a new function to replace the 'isActiveUserHOD()' function (from 7th August 2024) as we have now created an 'WhoApprovesMe' 
  # field in a new 'Usr_Approvals' table (user-level), that is better to check against.

  tmpCountAppearancesSQL = "SELECT COUNT(ID) FROM Usr_Approvals WHERE WhoApprovesMe = '{0}' OR WAM2 = '{0}' OR WAM3 = '{0}' OR WAM4 = '{0}'".format(userToCheck)
  tmpCountAppearances = runSQL(tmpCountAppearancesSQL, False, '', '')

  if int(tmpCountAppearances) > 0:
    return True
  else:
    return False

  # Above could be written as one line (but I prefer above for legibility):
  # return True if int(tmpCountAppearances) > 0 else False


def canApproveSelf(userToCheck):
  # This function will return boolean (True or False) to indicate whether the passed user can approve themselves (by checking if users email address is in the appover list)

  # get email address of user
  userEmail = runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(userToCheck), False, '', '')
  tmpHODemails = getUsersApproversEmail(forUser = userToCheck)

  if userEmail in tmpHODemails:
    return True
  else:
    return False


def canUserApproveFeeEarner(UserToCheck, FeeEarner):
  # This function will return boolean (True or False) to indicate whether the passed 'UserToCheck' can Approve the passed 'FeeEarner'

  # get email address of user
  userEmail = runSQL("SELECT EMailExternal FROM Users WHERE Code = '{0}'".format(UserToCheck), False, '', '')
  tmpHODemails = getUsersApproversEmail(forUser = FeeEarner)

  if userEmail in tmpHODemails:
    return True
  else:
    return False


def getUsersApproversEmail(forUser):
  # This function will return a list of email addresses of the passed forUser

  hodEmailSQL = "SELECT STRING_AGG(EMailExternal, '; ') FROM Users WHERE Code IN ("
  hodEmailSQL += "SELECT 'Who' = WhoApprovesMe FROM Usr_Approvals WHERE UserCode = '{0}' ".format(forUser)
  hodEmailSQL += "UNION SELECT 'Who' = WAM2 FROM Usr_Approvals WHERE UserCode = '{0}' ".format(forUser)
  hodEmailSQL += "UNION SELECT 'Who' = WAM3 FROM Usr_Approvals WHERE UserCode = '{0}' ".format(forUser)
  hodEmailSQL += "UNION SELECT 'Who' = WAM4 FROM Usr_Approvals WHERE UserCode = '{0}')".format(forUser)
  hodEmail = runSQL(hodEmailSQL, False, '', '')
  #hodEmail = runSQL(hodEmailSQL, True, 'There was an error getting approval users email address...', 'DEBUGGING - getUsersApproversEmail')
  return hodEmail


def getSQLDate(varDate):
  #Converts the passed varDate into SQL version date (YYYY-MM-DD)

  newDate = ''
  tmpDate = ''
  tmpDay = ''
  tmpMonth = ''
  tmpYear = ''
  mySplit = []
  finalStr = ''
  canContinue = False

  # If passed value is of 'DateTime' then convert to string
  if isinstance(varDate, DateTime) == True:
    tmpDate = varDate.ToString()
    canContinue = True

  # else if already a string, assingn passed date directly into newDate 
  elif isinstance(varDate, str) == True:
    tmpDate = varDate                       #datetime.datetime(varDate) '1/1/2020'
    canContinue = True

  if canContinue == True:
    # now to strip out the time element
    mySplit = []
    mySplit = tmpDate.split(' ')
    newDate = mySplit[0]

    #MessageBox.Show('newDate is ' + newDate)
    mySplit = []

    if len(newDate) >= 8:
      mySplit = newDate.split('/')

      tmpDay = mySplit[0]             #newDate.strftime("%d")
      tmpMonth = mySplit[1]           #newDate.strftime("%m")
      tmpYear = mySplit[2]            #newDate.strftime("%Y")

      testStr = str(tmpYear) + '-' + str(tmpMonth) + '-' + str(tmpDay)
      finalStr = testStr

    return finalStr
  
### character/type conversion tools ###################################################
  
def get_FullEntityRef(shortRef): 
  # Returns the full (15 character) length of the passed 'shortRef' Entity Reference
  if len(shortRef) < 4:
    myFinalString = shortRef
  else:
    leftPart = shortRef[0:3] 
    rightPart = shortRef[3:7]
    noZerosToAdd = 15 - len(shortRef)
    myZeros = ''

    for x in range(noZerosToAdd):
      myZeros += '0'

    # combine text elements to create full length ref
    myFullLenString = leftPart + myZeros + rightPart
    # now need to actually check if Entity exists - following returns 0 if not valid entity, else returns number of entities with that ref - should only be one)
    countOfEntities = 0
    countOfEntities = _tikitResolver.Resolve("[SQL: SELECT COUNT(Code) FROM Entities WHERE Code = '{0}']".format(myFullLenString))
    # if above count is zero, we return the short ref so other functions return 'error' otherwise we provide the full length code
    if int(countOfEntities) == 0:
      myFinalString = shortRef
    else:
      myFinalString = myFullLenString
 
    #MessageBox.Show("GetFullLenEntityRef - Input: " + str(shortRef) + "\nOutput: " + myFinalString + "\nCount of Entities: " + str(countOfEntities))
  return myFinalString

def getTextualTime(inputMinutes):
  # This function takes the 'inputMinutes' and returns a nicer string showing time including 'days'
  # Eg: if 'inputMinutes' = 2880, output will the '2 days + 00:00' (HH:MM)
  if inputMinutes == '':
    return '00:00'
  outputText = ''
  myMins = inputMinutes

  myHoursI = int(myMins/60)                         # This gives us an INTEGER (no deciaml) of num of minutes divided by 60
  minsRemainder = myMins % 60                       # This will give us just the REMAINDER of the same calculation (using mod)
  MessageBox.Show("error in get textual time 2")
  if myHoursI > 24:                                 # If there's more than 24 hours
    myDays = int(myHoursI/24)                       # This gives us an INTEGER (no decimal) of num of hours divided by 24
    timRemain = myHoursI % 24                       # This will give us just the REMAINDER of the same calculation (using mod)

    outputText = str(myDays)                        # this and following line just output the number of days including the word 'day(s)'
    outputText += ' days + ' if myDays > 1 else ' day + '
  else:
    timRemain = myHoursI                            # num of hours is less than 24, so we'll just use our current 'hours' variable
  MessageBox.Show("error in get textual time 3")
  if len(str(timRemain)) == 1:                      # next we output the hours remaining (including a leading zero if only one character long
    outputText += '0' + str(timRemain)
  else:
    outputText += str(timRemain)
  MessageBox.Show("error in get textual time 4")
  outputText += ":"                                 # add the hours and minutes separator
  if len(str(minsRemainder)) == 1:                  # similar to 'hours' above, but for 'minutes' (include leading zero if only one character long)
    outputText += '0' + str(minsRemainder)
  else:
    outputText += str(minsRemainder)

  return str(outputText)                            # finally return our 'outputText' to calling procedure

