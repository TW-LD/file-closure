<?xml version="1.0" encoding="utf-8"?>
<tfb>
  <fileclosure>
    <Init>
      <![CDATA[
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
from System.Windows.Data import Binding, IValueConverter
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import Brush, Brushes
from TWUtils import *

 # TODO: draft tables and structure for the checklist.

#Global Variables
UserIsHOD = False
all_ticked = False

## Main WIP Review DataGrid
class WIPreview(object):
  def __init__(self, myTicked, myRef, myClient, myMatDesc, myEntRef, myMatNo, myFENote, myTimeInactive, myLastUpdated, myLastActivity):
    self.iTemTicked       = myTicked
    self.wOurRef          = myRef
    self.wClientName      = myClient
    self.wMatDesc         = myMatDesc
    self.wEntRef          = myEntRef
    self.wMatNo           = myMatNo
    self.wFENote          = myFENote
    self.wTimeInactive    = myTimeInactive
    self.wTimeInactiveYMD = convert_days(number_of_days=myTimeInactive)
    self.wLastUpdated     = myLastUpdated
    self.wLastActivity = myLastActivity
    return
    
  def __getitem__(self, index):
    if index == 'EntityRef':
      return self.wEntRef
    elif index == 'MatterNo':
      return self.wMatNo
    elif index == 'TickedStatus':
      return self.iTemTicked
    elif index == 'OurRef':
      return self.wOurRef
    elif index == 'MatDesc':
      return self.wMatDesc
    elif index == 'FENote':           
      return self.wFENote
    elif index == 'TimeInactive':
      return self.wTimeInactive
    elif index == 'LastUpdated':
      return self.wLastUpdated
    elif index == 'ClientName':
      return self.wClientName
    elif index == 'TimeInactiveYMD':
      return self.wTimeInactiveYMD
    return None

# Conditional highlighting tool
class TimeInactiveToBrushConverter(IValueConverter):
    def Convert(self, value, targetType, parameter, culture):
        if not value:
            return Brushes.Transparent
        try:
            weeksInactive = int(value)
            if weeksInactive > 20:
                return Brushes.Red
            elif weeksInactive > 10:
                return Brushes.Orange
            else:
                return Brushes.LightGreen
        except:
            return Brushes.Transparent

    def ConvertBack(self, value, targetType, parameter, culture):
        raise NotImplementedError("No need for ConvertBack")

def refreshWIPReviewDataGrid(s, event):
  if cbo_FeeEarner.SelectedIndex == -1:
    MessageBox.Show("No items to show as a Fee Earner hasn't been selected from the drop-down.\n\nPlease note that only the Accounts department are able to select a different Fee Earner", "Refresh WIP Review List...")
    return

  # TODO: We need to redo the SQL below to grab the following fields: Our ref, Client Name, Matter Description, Time inactive, Archiving issues.
  # TODO: Need to add a table and SQL grab for fee earner notes for archiving.
  # TODO: MP: Changed to 'DaysInactive' (was initially 'week')

  wip_SQL = """
  SELECT
      -- Existing columns
      '0-OurRef'        = E.ShortCode + '/' + CONVERT(VARCHAR, M.Number),
      '1-Client Name'   = E.Name,
      '2-Matter Desc'   = M.Description,
      '3-EntRef'        = M.EntityRef,
      '4-MatNo'         = M.Number,
      '5-ArchivingNote' = FCH.ArchivingNote,
      '6-DaysInactive' = 
       DATEDIFF(DAY,(
          SELECT MAX(d)
          FROM
          (
              -- 1) Last Bill Date
              SELECT M.LastBillPostingDate AS d

              UNION ALL

              -- 2) Last Time Posting
              SELECT M.LastTimePostingDate AS d

              UNION ALL

              -- 3) Last Document (Case Manager) Date
              SELECT 
              (
                  SELECT TOP 1 CM.StepCreated
                  FROM View_CaseManagerMP CM
                  WHERE CM.EntityRef = M.EntityRef
                    AND CM.MatterRef = M.Number
                  ORDER BY CM.StepCreated DESC
              ) AS d
          ) AS AllDates
      ), GETDATE()), 
      '7-LastUpdated' = FCH.LastUpdated, 
	    '8-LastActivity' = (SELECT TOP 1 Narr FROM
          (
              -- 1) Last Bill Date
              SELECT 'Narr' = 'Last Bill Posting (' + CONVERT(VARCHAR(10), M.LastBillPostingDate, 103) + ')', 
					'myDate' = M.LastBillPostingDate 

              UNION ALL

              -- 2) Last Time Posting
              SELECT 'Narr' = 'Last Time Posting (' + CONVERT(VARCHAR(10), M.LastTimePostingDate, 103) + ')', 
					'myDate' = M.LastTimePostingDate

              UNION ALL

              -- 3) Last Document (Case Manager) Date
              SELECT TOP 1 'Narr' = 'Last document added to case (' + CONVERT(VARCHAR(10), CM.StepCreated, 103) + ')', 
						'myDate' = CM.StepCreated
						FROM View_CaseManagerMP CM
		                WHERE CM.EntityRef = M.EntityRef AND CM.MatterRef = M.Number
						ORDER BY CM.StepCreated DESC
          ) AS AllDates ORDER BY myDate DESC)

  FROM Matters M
      LEFT OUTER JOIN Entities E
          ON M.EntityRef = E.Code
      LEFT OUTER JOIN Usr_FileClosureHeader FCH
          ON FCH.EntityRef = M.EntityRef
          AND FCH.MatterNo = M.Number
  WHERE
      M.FeeEarnerRef = '{0}'
  ORDER BY [6-DaysInactive] DESC;
    """.format(cbo_FeeEarner.SelectedItem['Code'])

  _tikitDbAccess.Open(wip_SQL)
  mItem = []

  if _tikitDbAccess._dr is not None:
      dr = _tikitDbAccess._dr
      if dr.HasRows:
          while dr.Read():
              iTicked  = False
              iRef     = '' if dr.IsDBNull(0) else dr.GetString(0)  # 0-OurRef
              iClient  = '' if dr.IsDBNull(1) else dr.GetString(1)  # 1-Client Name
              iMatDesc = '' if dr.IsDBNull(2) else dr.GetString(2)  # 2-Matter Desc
              iEntRef  = '' if dr.IsDBNull(3) else dr.GetString(3)  # 3-EntRef
              iMatNo   = 0  if dr.IsDBNull(4) else dr.GetValue(4)   # 4-MatNo
              iFENote  = '' if dr.IsDBNull(5) else dr.GetString(5)  # 5-ArchivingNote
              iTimeInactive = 0 if dr.IsDBNull(6) else dr.GetValue(6)  # 6-TimeInactive
              iLastUpdated = '' if dr.IsDBNull(7) else dr.GetValue(7)  # 7-LastUpdated
              iLastActivity = '' if dr.IsDBNull(8) else dr.GetString(8)

              wip_item = WIPreview(
                  iTicked,
                  iRef,
                  iClient,
                  iMatDesc,
                  iEntRef,
                  iMatNo,
                  iFENote,
                  iTimeInactive,
                  iLastUpdated,
                  iLastActivity)
              mItem.append(wip_item)
      else:
          mItem.append(WIPreview(False, "-N/A-", "-No Data-", "-No Data-", "", 0, 0, 0, "", ""))
      dr.Close()
  else:
      mItem.append(WIPreview(False, "-N/A-", "-No Data-", "-No Data-", "", 0, 0, 0, "", ""))

  _tikitDbAccess.Close

  dg_WIPReview.ItemsSource = mItem

  return


def cellEdit_Finished(s, event):
  
  # Get column name
  tmpCol = event.Column
  tmpColName = tmpCol.Header    
  newDate = getSQLDate(_tikitResolver.Resolve("[SQL: SELECT GETDATE()]"))
  tmpEntity = dg_WIPReview.SelectedItem['EntityRef']
  tmpMatter = dg_WIPReview.SelectedItem['MatterNo']
  tmpNote = str(dg_WIPReview.SelectedItem['FENote'])
  updateSQL = 'UPDATE Usr_FileClosureHeader SET '
  countToUpdate = 0
  global UserIsHOD

  # count if there are any rows in Usr_AccWIP and if zero, add a new row with default data
  countExistingRows = _tikitResolver.Resolve("[SQL: SELECT COUNT(ID) FROM Usr_FileClosureHeader WHERE EntityRef = '{0}' AND MatterNo = {1}]".format(tmpEntity, tmpMatter))
  if int(countExistingRows) == 0:
    _tikitResolver.Resolve("[SQL: INSERT INTO Usr_FileClosureHeader (EntityRef, MatterNo, ArchivingNote) VALUES ('{0}', {1}, 'N')]".format(tmpEntity, tmpMatter))

  # get ID of row in Usr_AccWIP table
  IDtoUpdate = _tikitResolver.Resolve("[SQL: SELECT ID FROM Usr_FileClosureHeader WHERE EntityRef = '{0}' AND MatterNo = {1}]".format(tmpEntity, tmpMatter))

  # if name of column is 'Archiving Notes'
  if tmpColName == 'Archiving Notes':
    if str(dg_WIPReview.SelectedItem['FENote']) != lbl_tmpNote.Content:
      updateSQL += "ArchivingNote = '{0}', LastUpdated = '{1}' ".format(tmpNote.replace("'","''"), newDate)
      countToUpdate += 1

  if countToUpdate > 0:
    #Add Where
    updateSQL += "WHERE ID = {0}".format(IDtoUpdate)
    _tikitResolver.Resolve("[SQL: {0}]".format(updateSQL))
   
    countFEreply = _tikitResolver.Resolve("[SQL: SELECT COUNT(ID) FROM Usr_FileClosureHeader WHERE UserCode = '{0}']".format(_tikitUser))
    if int(countFEreply) == 0:
      _tikitResolver.Resolve("[SQL: INSERT INTO Usr_FileClosureHeader (UserCode, LastUpdated, EntityRef, MatterNo) VALUES ('{0}', '{1}', '{entity}', '{matter}')]".format(_tikitUser, newDate, entity=tmpEntity, matter=tmpMatter))
    else:
      _tikitResolver.Resolve("[SQL: UPDATE Usr_FileClosureHeader SET LastUpdated = '{0}' WHERE UserCode = '{1}']".format(newDate, _tikitUser))

  # just update the ticked counter
  return


class SortByList(object):
  def __init__(self, myCode, myName, mySQLCode):
    self.Ref = myCode
    self.Name = myName
    self.SQLCode = mySQLCode
    return
    
  def __getitem__(self, index):
    if index == 'Ref':
      return self.Code
    elif index == 'Name':
      return self.Name
    elif index == 'SQLCode':
      return self.SQLCode

def populate_SortByList(s, event):
  myItems = []
  
  myItems.append(SortByList(0, 'Our Ref', '[0-OurRef]'))
  myItems.append(SortByList(1, 'Our Ref Descending', '[0-OurRef] DESC'))
  cbo_SortBy.ItemsSource = myItems
  return

class UsersList(object):
  def __init__(self, myFECode, myFEName):
    self.Code = myFECode
    self.Name = myFEName
    return

  def __getitem__(self, index):
    if index == 'Code':
      return self.Code
    elif index == 'Name':
      return self.Name

def populate_FeeEarnersList(s, event): 
  mySQL = "SELECT Code, FullName FROM Users WHERE FeeEarner = 1 AND Locked = 0 AND UserStatus = 0 ORDER BY FullName"

  _tikitDbAccess.Open(mySQL)
  myFEitems = []

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          myCode = '-' if dr.IsDBNull(0) else dr.GetString(0)
          myName = '-' if dr.IsDBNull(1) else dr.GetString(1)

          myFEitems.append(UsersList(myCode, myName))  
    else:
      myFEitems.append(UsersList('-', '-'))

    dr.Close()
  _tikitDbAccess.Close()

  cbo_FeeEarner.ItemsSource = myFEitems
  return


def setFE_toCurrentUser(s, event):
  tCount = -1
  tMatchFound = False

  for xRow in cbo_FeeEarner.Items:
    tCount += 1
    if xRow.Code == _tikitUser:
      cbo_FeeEarner.SelectedIndex = tCount
      tMatchFound = True
      break

  if tMatchFound == True:
    cbo_SortBy.SelectedIndex = 5
  return

def totalNoOfItems():
  tmpCount = 0

  for row in dg_WIPReview.Items:
    tmpCount += 1
  return tmpCount

def totalTicked():
  tmpCount = 0

  for row in dg_WIPReview.Items:
    if row.iTemTicked == True:
      tmpCount += 1
  return tmpCount



def set_cboFE_Enabled(s, event):
  # first lookup user group ref 
  # 1 - Partner
  # 19 - Equity Partner
  # 8 - Accounts
  # 11 - System Admin
  isSysAdminOrAccountsOrPartner = _tikitResolver.Resolve("[SQL: SELECT COUNT(Code) FROM Users WHERE (UsertypeRef IN (8, 11) OR Partner = 1) AND Code = '" + str(_tikitUser) + "']")

  if int(isSysAdminOrAccountsOrPartner) == 1:
    cbo_FeeEarner.IsEnabled = True
  else:
    cbo_FeeEarner.IsEnabled = False
  return

# # # # # # # # #   T I M E   B R E A K D O W N   # # # # # # # # # # 
class matterTime(object):
  def __init__(self, myDate, myQtyOfT, myValOfT, myDesc, myNarr, myOrig):
    self.wipTDate = myDate
    self.wipQofT = myQtyOfT
    self.wipValOfTime = myValOfT
    self.wipDesc = myDesc
    self.wipNarr = myNarr
    self.wipOrig = myOrig
    return


def refresh_Matter_UnbilledTime(s, event):

  if dg_WIPReview.SelectedIndex == -1:
    ti_UnbilledTime.Header = 'WIP (Unbilled Time) - No data'
    return

  tmpEntity = dg_WIPReview.SelectedItem['EntityRef']
  tmpMatter = dg_WIPReview.SelectedItem['MatterNo']
  if dg_TimeUsers.SelectedIndex == -1:
    dg_TimeUsers.SelectedIndex = 0
  showFor = dg_TimeUsers.SelectedItem['Code']
  tmpCountRows = 0

  time_SQL = """SELECT '0-Transaction Date' = TT.TransactionDate, 
                '1-Quantity Of Time' = TT.QuantityOfTime / 60,
                '2-Value Of Time' = TT.ValueOfTime - TT.TimeValueBilled, 
                '3-Activity Type' = ActT.Description,
                '4-Narrative' = TT.Narratives + TT.Narrative2 + TT.Narrative3,
                '5-Originator' = (SELECT FullName FROM Users U WHERE U.Code = TT.Originator)
                FROM TimeTransactions TT, ActivityTypes ActT
                WHERE TT.ActivityCodeRef = ActT.Code AND ActT.ChargeType = 'C' AND TT.TransactionType LIKE '%Time'
                AND TT.EntityRef = '{entity}' AND 
                TT.MatterNoRef = {matter} AND TT.ValueOfTime - TT.TimeValueBilled > 0 """.format(entity=tmpEntity, matter=tmpMatter)
  if str(showFor) != 'x':
    time_SQL += "AND TT.Originator = '{0}' ".format(showFor)
  time_SQL += "ORDER BY TT.TransactionDate DESC"

  # Open and store items in code
  _tikitDbAccess.Open(time_SQL)
  tItem = []

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        iTDate = 0 if dr.IsDBNull(0) else dr.GetValue(0)  
        iQtyTime = 0 if dr.IsDBNull(1) else dr.GetValue(1)
        iValTime = 0 if dr.IsDBNull(2) else dr.GetValue(2)
        iActType = '' if dr.IsDBNull(3) else dr.GetString(3)
        iNarr = '' if dr.IsDBNull(4) else dr.GetString(4)
        iOrig = '' if dr.IsDBNull(5) else dr.GetString(5)
        nQtyTime = getTextualTime(iQtyTime)
        nValTime = '{:,.2f}'.format(iValTime)

        tItem.append(matterTime(iTDate, nQtyTime, nValTime, iActType, iNarr, iOrig))
        tmpCountRows += 1
    else:
      tItem.append(matterTime('', 0, 0.00, '-No Data-', '-No Data for ' + dg_WIPReview.SelectedItem['OurRef'] + '-', ''))

    dr.Close()
  else:
    tItem.append(matterTime('', 0, 0.00, '-No Data-', '-No Data-', ''))
  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_UnbilledTime.ItemsSource = tItem
  ti_UnbilledTime.Header = "WIP (Unbilled Time) - {0} entries...".format(tmpCountRows) 
  return


class matterDisbs(object):
  def __init__(self, myTNo, myRef, myDate, myNarr, myNet, myNetBilled, myNetPaid, myUnbilledVal, myNoDaysOS):
    self.disbTransNo = myTNo
    self.disbRef = myRef
    self.disbDate = myDate
    self.disbNarr = myNarr
    self.disbNet = myNet
    self.disbNetBilled = myNetBilled
    self.disbNetPaid = myNetPaid
    self.disbUnbilledVal = myUnbilledVal
    self.disbDaysOS = myNoDaysOS
    return

def refresh_Matter_UnbilledDisbs(s, event):

  if dg_WIPReview.SelectedIndex == -1:
    ti_UnbilledDisbs.Header = 'Unbilled Disbursements - No data'
    lbl_NO_DISBS.Visibility = Visibility.Visible
    dg_UnbilledDisbs.Visibility = Visibility.Hidden
    return

  tmpEntity = dg_WIPReview.SelectedItem['EntityRef']
  tmpMatter = dg_WIPReview.SelectedItem['MatterNo']
  tmpCount = 0

  disb_SQL = """
  SELECT 
      '0-Transaction Number' = AcD.TransactionNo, 
      '1-Disb Reference' = AcD.Ref, 
      '2-Date' = AcD.DisbDate, 
      '3-Narrative' = AcD.Narrative1, 
      '4-Net' = AcDV.Net, 
      '5-Net Billed' = AcDV.Net_Billed, 
      '6-Net Paid' = AcDV.Net_Paid, 
      '7-Unbilled Value' = AcDV.Net_Paid - AcDV.Net_Billed, 
      '8-Days Outstanding' = DATEDIFF(dd, RIGHT(AcD.DisbDate, 4) + SUBSTRING(AcD.DisbDate, 4, 2) + LEFT(AcD.DisbDate, 2), GETDATE())
  FROM 
      Ac_Disbursements AcD 
      JOIN Ac_Disbursements_VAT AcDV 
          ON AcD.TransactionNo = AcDV.Disbs_ID 
          AND AcD.AnticipatedNo = AcDV.AnticipatedNo
  WHERE 
      AcD.Type = 'UnBilled'  
      AND LEN(DisbDate) = 10 
      AND AcD.ClientCode = '{entity}' 
      AND AcD.MatterNo = {matter}
  ORDER BY 
      RIGHT(DisbDate, 4) + SUBSTRING(DisbDate, 4, 2) + LEFT(DisbDate, 2)
  """.format(entity=tmpEntity, matter=tmpMatter)

  # Open and store items in code
  _tikitDbAccess.Open(disb_SQL)
  dItem = []

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        iTNo = 0 if dr.IsDBNull(0) else dr.GetValue(0)  
        iRef = '' if dr.IsDBNull(1) else dr.GetString(1)
        iTDate = '' if dr.IsDBNull(2) else dr.GetString(2)
        iNarr = '' if dr.IsDBNull(3) else dr.GetString(3)
        iNet = 0 if dr.IsDBNull(4) else dr.GetValue(4)
        iNetBilled = 0 if dr.IsDBNull(5) else dr.GetValue(5)
        iNetPaid = 0 if dr.IsDBNull(6) else dr.GetValue(6)
        iUnbilledVal = 0 if dr.IsDBNull(7) else dr.GetValue(7)
        iDaysOS = 0 if dr.IsDBNull(8) else dr.GetValue(8)

        dItem.append(matterDisbs(iTNo, iRef, iTDate, iNarr, iNet, iNetBilled, iNetPaid, iUnbilledVal, iDaysOS))
        tmpCount += 1

    dr.Close()
  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_UnbilledDisbs.ItemsSource = dItem
  ti_UnbilledDisbs.Header = "Unbilled Disbursements ({0})".format(tmpCount)

  if dg_UnbilledDisbs.Items.Count == 0:
    lbl_NO_DISBS.Visibility = Visibility.Visible
    dg_UnbilledDisbs.Visibility = Visibility.Hidden
  else:
    lbl_NO_DISBS.Visibility = Visibility.Hidden
    dg_UnbilledDisbs.Visibility = Visibility.Visible
  return


class matterBills(object):
  def __init__(self, myOSTotal, myBillDate, myType, myPeriod, myYear, myDaysOS):
    self.ubOSTotal = myOSTotal
    self.ubBillDate = myBillDate
    self.ubType = myType
    self.ubPeriod = myPeriod
    self.ubYear = myYear
    self.ubDaysOS = myDaysOS
    return

def refresh_Matter_UnpaidBills(s, event):
  
  if dg_WIPReview.SelectedIndex == -1:
    ti_UnpaidBills.Header = "Unpaid Bills - No data"
    lbl_NO_BILLS.Visibility = Visibility.Visible
    dg_UnpaidBills.Visibility = Visibility.Hidden
    return

  tmpEntity = dg_WIPReview.SelectedItem['EntityRef']
  tmpMatter = dg_WIPReview.SelectedItem['MatterNo']
  tmpCount = 0

  bill_SQL = """
  SELECT '0-Outstanding Total' = AcBB.OutstandingTotal, 
        '1-Bill Date' = AcBB.BillDate, 
        '2-Type' = AcBB.Type, 
        '3-Period' = Period, 
        '4-Year' = Year, 
        '5-Days Outstanding' = DATEDIFF(dd, AcBB.BillDate, GETDATE())
  FROM Ac_Billbook AcBB
  WHERE AcBB.Type = 'Posted' 
    AND AcBB.OutstandingTotal > 0
    AND AcBB.EntityRef = '{entity}' 
    AND AcBB.MatterRef = {matter}
  ORDER BY AcBB.BillDate
  """.format(entity=tmpEntity, matter=tmpMatter)

  # Open and store items in code
  _tikitDbAccess.Open(bill_SQL)
  bItem = []
  bNoData = False

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        iOSTotal = 0 if dr.IsDBNull(0) else dr.GetValue(0)  
        iBDate = 0 if dr.IsDBNull(1) else dr.GetValue(1)
        iType = '' if dr.IsDBNull(2) else dr.GetString(2)
        iPeriod = 0 if dr.IsDBNull(3) else dr.GetValue(3)
        iYear = 0 if dr.IsDBNull(4) else dr.GetValue(4)
        iDaysOS = 0 if dr.IsDBNull(5) else dr.GetValue(5)

        bItem.append(matterBills(iOSTotal, iBDate, iType, iPeriod, iYear, iDaysOS))
        tmpCount += 1
    else:
      bItem.append(matterBills(0, 0, '-No Data for ' + dg_WIPReview.SelectedItem['OurRef'] + '-', 0, 0, 0))
      bNoData = True
    dr.Close()
  else:
    bItem.append(matterBills(0, 0, '-No Data-', 0, 0, 0))
    bNoData = True

  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_UnpaidBills.ItemsSource = bItem
  ti_UnpaidBills.Header = 'Unpaid Bills (' + str(tmpCount) + ')'

  if bNoData == True:
    lbl_NO_BILLS.Visibility = Visibility.Visible
    dg_UnpaidBills.Visibility = Visibility.Hidden
  else:
    lbl_NO_BILLS.Visibility = Visibility.Hidden
    dg_UnpaidBills.Visibility = Visibility.Visible
  return


def update_Details_Datagrids(s, event):
  #! Linked to XAML control.event(s): dg_WIPReview.SelectionChanged
  # This function is called when the user clicks the 'View Details' checkbox or selects a different matter from the main WIP datagrid.

  # if nothing selected
  if dg_WIPReview.SelectedIndex == -1:
    # populate controls with 'null' (empty) strings
    lbl_tmpNote.Content = ''
    lbl_tmpWOReq.Content = ''
    lbl_tmpWOtype.Content = ''

    lbl_EntRef.Content = ''
    lbl_MatNo.Content = ''
    lbl_OurRef.Content = ''
    lbl_ClientName.Content = ''
    lbl_MatDesc.Content = ''

    grp_WIPReview.Header = 'Details for selected matter: (NO MATTER CURRENTLY SELECTED)'

  else:
    # something is selected - populate controls with selected matter details
    lbl_tmpNote.Content = dg_WIPReview.SelectedItem['FENote']
    grp_WIPReview.Header = 'Details for selected matter: {0} - {1}'.format(dg_WIPReview.SelectedItem['OurRef'], dg_WIPReview.SelectedItem['MatDesc'])
    dg_WIPReview.BeginEdit()

    # populate 'Matter Details' tab in case user selects to view Archive details
    lbl_EntRef.Content = dg_WIPReview.SelectedItem['EntityRef']
    lbl_MatNo.Content = dg_WIPReview.SelectedItem['MatterNo']
    lbl_OurRef.Content = dg_WIPReview.SelectedItem['OurRef']
    lbl_ClientName.Content = dg_WIPReview.SelectedItem['ClientName']
    lbl_MatDesc.Content = dg_WIPReview.SelectedItem['MatDesc']

#  # if 'viewDetails' toggle button is checked
#  # TODO: Wouldn't it be better to always have this visible?  Are there any reasons to hide it?
#  if chk_ViewDetails.IsChecked == True:
#    # update other datagrids and not fogetting to update the ticked status
#    dg_WIPReview.Height = 427
#    grp_WIPReview.Visibility = Visibility.Visible
#    populate_TimeUsersToShow(s, event)
#    refresh_Matter_UnbilledTime(s, event)
#    refresh_Matter_UnbilledDisbs(s, event)
#    refresh_Matter_UnpaidBills(s, event)
#    refresh_dgClientLedger(s, event)
#
#    populate_andSetTabVisibility_MatterArchiveDetails()
#  else:
#    # if not checked, hide the details and make the main datagrid bigger
#    grp_WIPReview.Visibility = Visibility.Hidden
#    dg_WIPReview.Height = 790

    populate_TimeUsersToShow(s, event)
    refresh_Matter_UnbilledTime(s, event)
    refresh_Matter_UnbilledDisbs(s, event)
    refresh_Matter_UnpaidBills(s, event)
    refresh_dgClientLedger(s, event)
    #refresh_CaseDocs(s, event)
  return


def toggle_ViewArchiveMatterDetails(s, event):

  if btn_ViewArchiveDetails.IsChecked == False:
    # unload DataGrids
    dg_OutstandingAppointments.ItemsSource = None
    dg_OutstandingTasks.ItemsSource = None
    dg_Undertakings.ItemsSource = None
    dg_ForwardPostedItems.ItemsSource = None
    dg_UnclearedBankRecItems.ItemsSource = None
    dg_UnprocessedSlips.ItemsSource = None
    dg_NonZeroBalances.ItemsSource = None
    dg_PostToReview.ItemsSource = None
    dg_CheckedOutDocuments.ItemsSource = None 
  else:
    populate_andSetTabVisibility_MatterArchiveDetails()
  return


def populate_andSetTabVisibility_MatterArchiveDetails():
  # populate the 'Matter Archive Details' tab and set visibility (hide if nothing in data grids)

  # call functions to populate DataGrids
  refresh_OutstandingAppointments()
  refresh_OutstandingTasks()
  refresh_UndertakingsList()
  refresh_forwardPostedItems()
  refresh_UnclearedBankReceivedItems()
  refresh_UnprocessedSlips()
  refresh_NonzeroBalances()
  refresh_PostToReview()
  refresh_CheckedOutDocuments()

  # now set visibility based on DataGrid.Items.Count
  if dg_OutstandingAppointments.Items.Count == 0:
    ti_OsAppts.Visibility = Visibility.Collapsed
  else:
    ti_OsAppts.Visibility = Visibility.Visible

  if dg_OutstandingTasks.Items.Count == 0:
    ti_OsTasks.Visibility = Visibility.Collapsed
  else:
    ti_OsTasks.Visibility = Visibility.Visible

  if dg_Undertakings.Items.Count == 0:
    ti_Undertaking.Visibility = Visibility.Collapsed
  else:
    ti_Undertaking.Visibility = Visibility.Visible

  if dg_ForwardPostedItems.Items.Count == 0:
    ti_ForPI.Visibility = Visibility.Collapsed
  else:
    ti_ForPI.Visibility = Visibility.Visible

  if dg_UnclearedBankRecItems.Items.Count == 0:
    ti_UnBankRec.Visibility = Visibility.Collapsed
  else:
    ti_UnBankRec.Visibility = Visibility.Visible

  if dg_UnprocessedSlips.Items.Count == 0:
    ti_UnpSlip.Visibility = Visibility.Collapsed
  else:
    ti_UnpSlip.Visibility = Visibility.Visible

  if int(NonZeroBalancesLabel.Content) == 0:
    ti_NZB.Visibility = Visibility.Collapsed
  else:
    ti_NZB.Visibility = Visibility.Visible

  if dg_PostToReview.Items.Count == 0:
    ti_PtR.Visibility = Visibility.Collapsed
  else:
    ti_PtR.Visibility = Visibility.Visible

  if dg_CheckedOutDocuments.Items.Count == 0:
    ti_COD.Visibility = Visibility.Collapsed
  else:
    ti_COD.Visibility = Visibility.Visible
  return 


class timeUsers(object):
  def __init__(self, myTick, myToShow, myTCode, myTotalTime, myValueOfTime):
    self.wipTTick = myTick
    self.wipTToShow = myToShow
    self.wipTCode = myTCode
    self.wipTtotalTime = myTotalTime
    self.wipTvalueOfTime = myValueOfTime
    return


  def __getitem__(self, index):
    if index == 'Tick':
      return self.wipTTick
    elif index == 'To Show':
      return self.wipTToShow
    elif index == 'Code':
      return self.wipTCode
    elif index == 'Total Time':
      return self.wipTtotalTime
    elif index == 'Value of Time':
      return self.wipTvalueOfTime

def populate_TimeUsersToShow(s, event):
  if dg_WIPReview.SelectedIndex == -1:
    return

  # add default items
  myItem = []
  tCount = 0
  myEntRef = dg_WIPReview.SelectedItem['EntityRef']
  myMatNo = str(dg_WIPReview.SelectedItem['MatterNo'])
  
  # now add all other users
  time_SQL = """
  SELECT ToShowName, Code, TotalTime, ValOfTime 
  FROM (
      SELECT 'ToShowName' = 'SHOW ALL TIME', 
            'Code' = 'x', 
            'TotalTime' = SUM(TT.QuantityOfTime) / 60, 
            'ValOfTime' = ISNULL(SUM(TT.ValueOfTime) - SUM(TT.TimeValueBilled), 0.00), 
            'mOrder' = 0 
      FROM TimeTransactions TT 
      WHERE TT.EntityRef = '{entity}' 
        AND TT.MatterNoRef = {matter} 
        AND TT.ValueOfTime - TT.TimeValueBilled > 0

      UNION 

      SELECT 'ToShowName' = U.FullName, 
            'Code' = U.Code, 
            'TotalTime' = SUM(TT.QuantityOfTime) / 60, 
            'ValOfTime' = ISNULL(SUM(TT.ValueOfTime) - SUM(TT.TimeValueBilled), 0.00), 
            'mOrder' = 2 
      FROM TimeTransactions TT 
      JOIN Users U ON TT.Originator = U.Code 
      WHERE TT.EntityRef = '{entity}' 
        AND TT.MatterNoRef = {matter} 
        AND TT.ValueOfTime - TT.TimeValueBilled > 0 
      GROUP BY U.FullName, U.Code
  ) AS tmpT 
  ORDER BY mOrder
  """.format(entity=myEntRef, matter=myMatNo)

  _tikitDbAccess.Open(time_SQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        tCount += 1
        iTick = True if tCount == 1 else False  
        iToShow = '' if dr.IsDBNull(0) else dr.GetString(0)
        iCode = '' if dr.IsDBNull(1) else dr.GetString(1)
        iTotalTime = 0 if dr.IsDBNull(2) else dr.GetValue(2)
        iValOfTime = 0 if dr.IsDBNull(3) else dr.GetValue(3)  
        nTotalTime = getTextualTime(iTotalTime)

        nValOfTime = '{:,.2f}'.format(iValOfTime)
        
        myItem.append(timeUsers(iTick, iToShow, iCode, nTotalTime, nValOfTime))

      dr.Close()
  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_TimeUsers.ItemsSource = myItem
  return


def getTextualTime(inputMinutes):
  # This function takes the 'inputMinutes' and returns a nicer string showing time including 'days'
  # Eg: if 'inputMinutes' = 2880, output will the '2 days' (HH:MM)
  # eg: if 'input' = 1440, returns '1 day'
  # ...   if = 3260, returns: '2 days + 06:20'
  # P4W tends to just increment 'hours' past 24 which jars my brain, so one then needs to divide by 24 to get number of days etc. This function simplifies by giving us the day count and then shows the remainder as 'valid' time units

  # Calculate how many whole days are in the total minutes
  days = inputMinutes // (60 * 24)

  # Get the remaining minutes once days are extracted
  remaining_minutes = inputMinutes % (60 * 24)

  # Calculate hours and minutes from what's left
  hours = remaining_minutes // 60
  minutes = remaining_minutes % 60

  # Construct the output string
  if days > 1:
    if hours == 0 and minutes == 0:
      #return f"{days} days"
      # without 'f strings'
      return "{0} days".format(days)
    else:
      #return f"{days} days + {hours:02}:{minutes:02}"
      # without 'f strings'
      return "{0} days + {1:02}:{2:02}".format(days, hours, minutes)
  elif days == 1:
    #return f"1 day + {hours:02}:{minutes:02}"
    if hours == 0 and minutes == 0:
      return "1 day"
    else:
      #return f"1 day + {hours:02}:{minutes:02}"
      # without 'f strings'
      return "1 day + {0:02}:{1:02}".format(hours, minutes)
      
  else:
    #return f"{hours:02}:{minutes:02}"
    # without 'f strings'
    return "{0:02}:{1:02}".format(hours, minutes)


def btn_Help_Clicked(s, event):
  kbLink = "https://thackraywilliams764.workplace.com/work/knowledge/2748766381952851" 
  System.Diagnostics.Process.Start(r'{}'.format(kbLink))  
  return

# #  C L I E N T   L E D G E R   -   C L A S S   A N D   R E F R E S H  #  #
class cls_dgClientLedger(object):
  def __init__(self, myclDate, myclRef, myclType, myclNarr, myclVATo, myclDebitO, myclCreditO, myclBalO, myclDebitC, myclCreditC, myclBalC, myclUnbilledDisb, myclUnpaidBill):

    self.clDate = myclDate
    self.clRef = myclRef
    self.clType = myclType
    self.clNarr = myclNarr
    self.clVATo = myclVATo
    self.clDebitO = myclDebitO
    self.clCreditO = myclCreditO
    self.clBalO = myclBalO
    self.clDebitC = myclDebitC
    self.clCreditC = myclCreditC
    self.clBalC = myclBalC
    if myclUnbilledDisb == 'Y':
      self.clUnbilledDisb = True
    else:
      self.clUnbilledDisb = False
    if myclUnpaidBill == 'Y':
      self.clUnpaidBill = True
    else:
      self.clUnpaidBill = False
    return


  def __getitem__(self, index): 

    if index == 'clRef':
      return self.clRef
    elif index == 'clType':
      return self.clType
    elif index == 'clNarr':
      return self.clNarr
    elif index == 'clVATo':
      return self.clVATo
    elif index == 'clDebitO':
      return self.clDebitO
    elif index == 'clCreditO':
      return self.clCreditO
    elif index == 'clBalO':
      return self.clBalO
    elif index == 'clDebitC':
      return self.clDebitC
    elif index == 'clCreditC':
      return self.clCreditC
    elif index == 'clBalC':
      return self.clBalC
    elif index == 'clUnbilledDisb':
      return self.clUnbilledDisb
    elif index == 'clUnpaidBill':
      return self.clUnpaidBill


def refresh_dgClientLedger(s, event):
  
  if dg_WIPReview.SelectedIndex == -1:
    ti_ClientLedger.Header = 'Client Ledger - No data'
    return

  myEntRef = dg_WIPReview.SelectedItem['EntityRef']
  myMatNo = str(dg_WIPReview.SelectedItem['MatterNo'])
  tmpCount = 0

  dgClientLedgerSQL = """
  SELECT 
      'Date' = CONVERT(VARCHAR(10), myCL.Date, 103), 
      myCL.Ref, 
      myCL.Type, 
      myCL.Narrative, 
      'VAT - O' = myCL.VAT, 
      'Debit - O' = myCL.Debit, 
      'Credit - O' = myCL.Credit, 
      'Balance - O' = SUM(myCL.RowTotal) OVER (ORDER BY myCL.ID), 
      'Debit - C' = myCL.[Debit - C], 
      'Credit - C' = myCL.[Credit - C], 
      'Balance - C' = SUM(myCL.RowTotalC) OVER (ORDER BY myCL.ID), 
      myCL.[Unbilled Disb], 
      myCL.[Unpaid Bill]
  FROM (
      SELECT TOP(10000) 
          ID, 
          'Date' = PostedDate, 
          'Ref' = Ref, 
          'Type' = Posting_Type, 
          'Narrative' = Narrative1 + Narrative2 + Narrative3, 
          'VAT' = VAT_In + VAT_Out, 
          'Debit' = Disbursements + Costs + Expenses - Office_Credit, 
          'Credit' = Office_Credit, 
          'RowTotal' = CASE 
                          WHEN Office_Credit > 0 THEN Office_Credit - (2 * Office_Credit)
                          WHEN UnBilledDisbursement = 1 THEN Disbursements + Costs + Expenses
                          ELSE Disbursements + Costs + Expenses + VAT_In + VAT_Out 
                      END, 
          'Debit - C' = Client_Debit, 
          'Credit - C' = Client_Credit, 
          'RowTotalC' = Client_Credit - Client_Debit, 
          'Unbilled Disb' = CASE WHEN UnBilledDisbursement = 1 THEN 'Y' ELSE 'N' END, 
          'Unpaid Bill' = CASE WHEN UnPaidBill = 1 THEN 'Y' ELSE 'N' END 
      FROM Ac_Client_Ledger_Transactions 
      WHERE Client_Code = '{entity}' AND Matter_No = {matter} 
      ORDER BY PostedDate
  ) AS myCL
  """.format(entity=myEntRef, matter=myMatNo)

  _tikitDbAccess.Open(dgClientLedgerSQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    myItem = []
    if dr.HasRows:
      while dr.Read():
        aDate = '' if dr.IsDBNull(0) else dr.GetString(0)
        aRef = '' if dr.IsDBNull(1) else dr.GetString(1)
        aType = '' if dr.IsDBNull(2) else dr.GetString(2)
        aNarr = '' if dr.IsDBNull(3) else dr.GetString(3)
        aVATo = '' if dr.IsDBNull(4) else dr.GetValue(4)
        aDbtO = '' if dr.IsDBNull(5) else dr.GetValue(5)
        aCdtO = '' if dr.IsDBNull(6) else dr.GetValue(6)
        aBalO = '' if dr.IsDBNull(7) else dr.GetValue(7)
        aDbtC = '' if dr.IsDBNull(8) else dr.GetValue(8)
        aCdtC = '' if dr.IsDBNull(9) else dr.GetValue(9)
        aBalC = '' if dr.IsDBNull(10) else dr.GetValue(10)
        aUND = '' if dr.IsDBNull(11) else dr.GetString(11)
        aUNB = '' if dr.IsDBNull(12) else dr.GetString(12)

        myItem.append(cls_dgClientLedger(aDate, aRef, aType, aNarr, aVATo, aDbtO, aCdtO, aBalO, aDbtC, aCdtC, aBalC, aUND, aUNB))
        tmpCount += 1

    dr.Close()
  _tikitDbAccess.Close

  dgClientLedger.ItemsSource = myItem
  if tmpCount == 0:
    ti_ClientLedger.Header = 'Client Ledger - No data'
  else:
    ti_ClientLedger.Header = 'Client Ledger - {0} items'.format(tmpCount)
  return
  
  
## E N D   O F   C L I E N T    L E D G E R    D A T A G R I D   F U N C T I O N S ##

def myOnFormLoadEvent(s, event):
  populate_SortByList(s, event)
  populate_FeeEarnersList(s, event)
  setFE_toCurrentUser(s, event)
  set_cboFE_Enabled(s, event)
  refreshWIPReviewDataGrid(s, event)
  update_Details_Datagrids(s, event)
  # determine if user is can Approve their own 'write-offs' (if they can, the following should return True, else False is returned)
  global UserIsHOD
  UserIsHOD = canApproveSelf(userToCheck = _tikitUser)
  ti_CaseDocs.Visibility = Visibility.Collapsed
  return

def cbo_FeeEarner_SelectionChanged(s, event):
  refreshWIPReviewDataGrid(s, event)

  return

# # # # # # # # # # # #  C A S E   D O C S   -   F U N C T I O N S  # # # # # # # # # # # # # # 
def CaseDoc_SelectionChanged(s, event):
  if dg_CaseManagerDocs.SelectedIndex == -1:
    btn_OpenCaseDoc.IsEnabled = False
  else:
    if dg_CaseManagerDocs.SelectedItem['Path'] != '':
      btn_OpenCaseDoc.IsEnabled = True
  return


def open_Selected_CaseDoc(s, event):
  tmpPath = dg_CaseManagerDocs.SelectedItem['Path']
  tmpName = dg_CaseManagerDocs.SelectedItem['Desc']

  if tmpPath == '':
    MessageBox.Show("There doesn't appear to be a path to this document: \n" + str(tmpName))
  else:
    System.Diagnostics.Process.Start(r'{}'.format(tmpPath))
  return


# Case Docs DataGrid
class CaseDocs(object):
  def __init__(self, mySID, mySDesc, mySCreated, mySPath, mySAgenda):
    self.sID = mySID
    self.sDescription = mySDesc
    self.sCreated = mySCreated
    self.sDocPath = mySPath
    self.sAgenda = mySAgenda
    return

  def __getitem__(self, index):
    if index == 'ID':
      return self.sID
    elif index == 'Desc':
      return self.sDescription
    elif index == 'Created':
      return self.sCreated 
    elif index == 'Path':
      return self.sDocPath 
    elif index == 'Agenda':
      return self.SAgenda


def refresh_CaseDocs(s, event):
  if dg_WIPReview.SelectedIndex == -1 or ti_CaseDocs.Visibility == Visibility.Collapsed:
    dg_CaseManagerDocs.ItemsSource = None
    return 

  sSQL = """SELECT CI.ItemID, CI.Description, CI.CreationDate,  CMS.FileName, Agenda.Description 
            FROM Cm_CaseItems CI 
            INNER JOIN Cm_Steps CMS ON CMS.ItemID = CI.ItemID 
            INNER JOIN Cm_Agendas CMA ON CI.ParentID = CMA.ItemID 
            INNER JOIN Cm_CaseItems Agenda ON CMA.ItemID = Agenda.ItemID 
            WHERE CMS.FileName <> '' """

  if cbo_AgendaName.SelectedIndex == 0:
    # show ALL documents for current matter
    sSQL += "AND CMA.EntityRef = '{0}' AND CMA.MatterNo = {1} ORDER BY Agenda.ItemID, CI.ItemOrder ".format(dg_WIPReview.SelectedItem['EntityRef'], dg_WIPReview.SelectedItem['MatterNo'])
    dg_CaseManagerDocs.Columns[0].Visibility = Visibility.Visible
  if cbo_AgendaName.SelectedIndex > 0:
    # just show docs for the selected Agenda
    tmpAgendaName = cbo_AgendaName.SelectedItem['ID']
    sSQL += "AND Agenda.ItemID = {0} ORDER BY CI.ItemOrder ".format(tmpAgendaName)
    dg_CaseManagerDocs.Columns[0].Visibility = Visibility.Hidden
  sItem = []

  # Open and store items in code
  _tikitDbAccess.Open(sSQL)

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        aID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
        aDesc = '' if dr.IsDBNull(1) else dr.GetString(1)
        aDate = 0 if dr.IsDBNull(2) else dr.GetValue(2)
        aPath = '' if dr.IsDBNull(3) else dr.GetString(3)
        aAgenda = '' if dr.IsDBNull(4) else dr.GetString(4)

        sItem.append(CaseDocs(aID, aDesc, aDate, aPath, aAgenda))
    dr.Close()
  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_CaseManagerDocs.ItemsSource = sItem
  lbl_OurRefCD.Content = dg_WIPReview.SelectedItem['OurRef']
  return


# Agenda Items
class cboAgendaNames(object):
  def __init__(self, myAgendaID, myAgendaName, myDefault):
    self.AgendaID = myAgendaID
    self.AgendaName = myAgendaName
    self.mIsDefault = myDefault

    if myAgendaName == 'Case History':
      self.mIsDefault = 1
    else:
      self.mIsDefault = myDefault
    return

  def __getitem__(self, index):
    if index == 'ID':
      return self.AgendaID
    elif index == 'Name':
      return self.AgendaName
    elif index == 'Default':
      return self.mIsDefault


def POPULATE_AGENDA_NAMES(s, event):
  if dg_WIPReview.SelectedIndex == -1 or ti_CaseDocs.Visibility == Visibility.Collapsed:
    return

  mySQL = """
  SELECT 
      Cm_CaseItems.Description, 
      Cm_Agendas.ItemID, 
      Cm_Agendas.Default_Agenda
  FROM 
      Cm_Agendas
      LEFT JOIN Cm_CaseItems ON Cm_Agendas.ItemID = Cm_CaseItems.ItemID
  WHERE 
      Cm_Agendas.EntityRef = '{entity_ref}' 
      AND Cm_Agendas.MatterNo = {matter_no}
  ORDER BY 
      Cm_CaseItems.Description
  """.format(
      entity_ref=dg_WIPReview.SelectedItem['EntityRef'], 
      matter_no=dg_WIPReview.SelectedItem['MatterNo'])

  _tikitDbAccess.Open(mySQL)
  itemA = []
  itemA.append(cboAgendaNames(0, '(ALL)', 0))

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          iAgendaName = '-' if dr.IsDBNull(0) else dr.GetString(0)
          iAgendaID = '-' if dr.IsDBNull(1) else dr.GetValue(1)
          iDefault = 0 if dr.IsDBNull(2) else dr.GetValue(2)
          itemA.append(cboAgendaNames(iAgendaID, iAgendaName, iDefault))
    dr.Close()
  _tikitDbAccess.Close()

  # Set set source of the Agenda Names combo box to the list of items we created
  cbo_AgendaName.ItemsSource = itemA
  if cbo_AgendaName.Items.Count > 0:
    cbo_AgendaName.SelectedIndex = 1

  return


def caseDocsPanel_refresh(s, event):
  POPULATE_AGENDA_NAMES(s, event)
  refresh_CaseDocs(s, event)
  return

# #  END: C A S E   D O C S   -   F U N C T I O N S  # #


def convert_days(number_of_days):
    """
    Convert a number of days into a human-readable string of years, weeks, and days.
    
    Args:
        number_of_days (int): The number of days to convert.
    
    Returns:
        str: A string representation in the form "x years, y weeks, z days"
             (omitting any zero units, except when all are zero).
    """
    # Calculate the number of years, months, and days
    years = number_of_days // 365
    remaining_days = number_of_days % 365
    weeks = remaining_days // 7
    days = remaining_days % 7

    parts = []
    # Only add the part if it's non-zero
    if years:
        parts.append("%d year%s" % (years, "s" if years != 1 else ""))
    if weeks:
        parts.append("%d week%s" % (weeks, "s" if weeks != 1 else ""))
    # Always include days if nothing else is present (e.g. 0 days)
    if days or not parts:
        parts.append("%d day%s" % (days, "s" if days != 1 else ""))

    # Join the parts with a comma and a space
    return ", ".join(parts)


# ---- OUTSTANDING APPOINTMENTS START -----
class OutstandingAppointmentsObject(object):
  def __init__(self, Type, User, ApptDate, ApptSubject, ApptType, CaseItemNo, CaseDesc, CaseAgenda):
    self.Type = Type
    self.User = User
    self.ApptDate = ApptDate
    self.ApptSubject = ApptSubject
    self.ApptType  = ApptType
    self.CaseItemNo = CaseItemNo
    self.CaseDesc = CaseDesc
    self.CaseAgenda = CaseAgenda
   
  def __getitem__(self, index):
    return None

def refresh_OutstandingAppointments():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False
  
  mySQL = """SELECT myATab.Type, myATab.Username, 'Appt Date' = myATab.DateStamp, 'Appt Desc' = myATab.Description, 'Appt Type' = myATab.Type, 'Case Item No' = myATab.CaseItemRef, 
          'Case Description' = CI.Description, 'Case Agenda' = AG.Description FROM (
          SELECT [EntityRef], [MatterNoRef], 'Type' = 'Individual', [Username], [DateStamp], [Description], [AppointmentTypes], [CaseItemRef] FROM Diary_Appointments WHERE DeleteThis = 0 
          UNION ALL 
          SELECT [EntityRef], [MatterNoRef], 'Type' = 'Group', [Username], [DateStamp], [Description], [AppointmentTypes], [CaseItemRef] FROM Diary_GroupAppointments WHERE DeleteThis = 0 
          ) as myATab 
          LEFT OUTER JOIN Cm_CaseItems CI ON CI.ItemID = myATab.CaseItemRef 
          LEFT OUTER JOIN Cm_CaseItems AG ON CI.ParentID = AG.ItemID 
          WHERE myATab.EntityRef = '{0}' AND myATab.MatterNoRef = {1}
          ORDER BY myATab.DateStamp""".format(lbl_EntRef.Content, lbl_MatNo.Content)
  
  _tikitDbAccess.Open(mySQL)
  items = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          Type = 'None' if dr.IsDBNull(0) else dr.GetString(0)
          User = 'None' if dr.IsDBNull(1) else dr.GetString(1)
          ApptDate = 'None' if dr.IsDBNull(2) else dr.GetValue(2)
          ApptSubject = 'None' if dr.IsDBNull(3) else dr.GetString(3)
          ApptType = 'None' if dr.IsDBNull(4) else dr.GetString(4)
          CaseItemNo = 'None' if dr.IsDBNull(5) else dr.GetValue(5)
          CaseDesc = 'None' if dr.IsDBNull(6) else dr.GetString(6)
          CaseAgenda = 'None' if dr.IsDBNull(7) else dr.GetString(7)

          items.append(OutstandingAppointmentsObject(Type, User, ApptDate, ApptSubject, ApptType, CaseItemNo, CaseDesc, CaseAgenda))
      
  dr.Close()
  _tikitDbAccess.Close()

  dg_OutstandingAppointments.ItemsSource = items
  OutstandingAppointmentsLabel.Content = str(dg_OutstandingAppointments.Items.Count)
  lbl_OSapptTask.Content = str(dg_OutstandingTasks.Items.Count + dg_OutstandingAppointments.Items.Count)
  return True
# ----- OUTSTANDING APPOINTMENTS END -----

# ----- OUTSTANDING TASKS START -----
class OutstandingTasksObject(object):
  def __init__(self, Type, User, TaskDate, TaskDesc, CaseItemNo, CaseDesc, CaseAgenda):
    self.Type = Type
    self.User = User
    self.TaskDate = TaskDate
    self.TaskDesc = TaskDesc
    self.CaseItemNo = CaseItemNo
    self.CaseDesc = CaseDesc
    self.CaseAgenda = CaseAgenda
   
  def __getitem__(self, index):
    return None

def refresh_OutstandingTasks():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False

  mySQL = """SELECT myTab.Type, myTab.Username, myTab.DateStamp, 'Task Desc' = myTab.Description, 'Case Item No' = myTab.CaseItemRef, 
            'Case Description' = CI.Description, 'Case Agenda' = AG.Description FROM (
            SELECT [EntityRef], [MatterNoRef], 'Type' = 'Individual', [Username], [DateStamp], [Description], [CaseItemRef] FROM Diary_Tasks WHERE DeleteThis = 0 
            UNION ALL 
            SELECT [EntityRef], [MatterNoRef], 'Type' = 'Group', [Username], [DateStamp], [Description], [CaseItemRef] FROM Diary_GroupTasks WHERE DeleteThis = 0) as myTab 
            LEFT OUTER JOIN Cm_CaseItems CI ON CI.ItemID = myTab.CaseItemRef 
            LEFT OUTER JOIN Cm_CaseItems AG ON CI.ParentID = AG.ItemID 
            WHERE myTab.EntityRef = '{0}' AND myTab.MatterNoRef = {1} ORDER BY myTab.DateStamp""".format(lbl_EntRef.Content, lbl_MatNo.Content)

  _tikitDbAccess.Open(mySQL)
  items = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          Type = 'None' if dr.IsDBNull(0) else dr.GetString(0)
          User = 'None' if dr.IsDBNull(1) else dr.GetString(1)
          TaskDate = 'None' if dr.IsDBNull(2) else dr.GetValue(2)
          TaskDesc = 'None' if dr.IsDBNull(3) else dr.GetString(3)
          CaseItemNo = 'None' if dr.IsDBNull(4) else dr.GetValue(4)
          CaseDesc = 'None' if dr.IsDBNull(5) else dr.GetString(5)
          CaseAgenda = 'None' if dr.IsDBNull(6) else dr.GetString(6)

          items.append(OutstandingTasksObject(Type, User, TaskDate, TaskDesc, CaseItemNo, CaseDesc, CaseAgenda))
    
  dr.Close()
  _tikitDbAccess.Close()

  dg_OutstandingTasks.ItemsSource = items
  OutstandingTasksLabel.Content = str(dg_OutstandingTasks.Items.Count)
  lbl_OSapptTask.Content = str(dg_OutstandingTasks.Items.Count + dg_OutstandingAppointments.Items.Count)
  return True
# ----- OUTSTANDING TASKS END -----

# ----- FORWARD POSTED ITEMS START -----
class ForwardPostedItemsObject(object):
  
  def __init__(self, Period, Year, Type, ClientBalance, OfficeBalance, UnpaidBillBalance, UnbilledDisbBalance, DepositedBalance, UnbilledTimeBalance, UnbilledTimeValue):
    self.Period = Period
    self.Year = Year
    self.Type = Type
    self.ClientBalance = ClientBalance
    self.OfficeBalance = OfficeBalance
    self.UnpaidBillBalance = UnpaidBillBalance
    self.UnbilledDisbBalance = UnbilledDisbBalance
    self.DepositedBalance = DepositedBalance
    self.UnbilledTimeBalance = UnbilledTimeBalance
    self.UnbilledTimeValue = UnbilledTimeValue
  
  def __getitem__(self, index):
    return None
 
def refresh_forwardPostedItems():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False

  mySQL = """SELECT [Period] = Period, 
              [Year] = Year, 
              [Type] = CASE WHEN UnbilledTimeBalance > 0 THEN 'Time' ELSE 'Financial' END, 
              [ClientBalance] = ClientAcBalance, 
              [OfficeBalance] = OfficeAcBalance, 
              [UnpaidBillBalance] = UnpaidBillBalance, 
              [UnbilledDisbBalance] = UnbilledDisbBalance, 
              [DepositAcBalance] = DepositAcBalance, 
              [UnbilledTimeBalance] = UnbilledTimeBalance, 
              [UnbilledTimeBalanceValue] = UnbilledTimeBalanceValue 
              FROM Ac_Forward_Matter_Balances 
              WHERE EntityRef = '{0}' AND MatterNo = {1}""".format(lbl_EntRef.Content, lbl_MatNo.Content)
    
  _tikitDbAccess.Open(mySQL)
  items = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          Period = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          Year = 0 if dr.IsDBNull(1) else dr.GetValue(1) 
          Type = '' if dr.IsDBNull(2) else dr.GetString(2) 
          ClientBalance = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          OfficeBalance = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          UnpaidBillBalance = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          UnbilledDisbBalance = 0 if dr.IsDBNull(6) else dr.GetValue(6)
          DepositedBalance = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          UnbilledTimeBalance = 0 if dr.IsDBNull(8) else dr.GetValue(8)
          UnbilledTimeValue = 0 if dr.IsDBNull(9) else dr.GetValue(9)
          items.append(ForwardPostedItemsObject(Period, Year, Type, ClientBalance, OfficeBalance, UnpaidBillBalance, UnbilledDisbBalance, DepositedBalance, UnbilledTimeBalance, UnbilledTimeValue))
          
  dr.Close()
  _tikitDbAccess.Close()

  dg_ForwardPostedItems.ItemsSource = items
  ForwardPostedItemsLabel.Content = str(dg_ForwardPostedItems.Items.Count)
  lbl_ForwardPostedItems.Content = str(dg_ForwardPostedItems.Items.Count)
  return True
# ----- FORWARD POSTED ITEMS END -----

# ---- UNCLEARED BANK RECEVIED ITEMS START -----
class UnclearedBankReceviedItemsObject(object):
  def __init__(self, Bank, TransNo, BatchNo, ChequeDate, Ref, Payee, Debit, Amount):
    self.Bank = Bank
    self.TransNo = TransNo
    self.BatchNo = BatchNo
    self.ChequeDate = ChequeDate
    self.Ref = Ref
    self.Payee = Payee
    self.Debit = Debit
    self.Amount = Amount
   
  def __getitem__(self, index):
    return None

def refresh_UnclearedBankReceivedItems():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False

  mySQL = """SELECT [Bank] = (Ac_Bank_Codes.Name + ' - ' + Branches.Description), 
              [Trans No] = CONVERT(VARCHAR(22),  Ac_Bank_Rec.Transaction_No), 
              [Batch No] = CONVERT(VARCHAR(22),  Ac_Bank_Rec.Batch_No), 
              [Cheque Date] = Ac_Bank_Rec.ChequeDate, 
              [Ref] = Ac_Bank_Rec.Ref, 
              [Payee] = Ac_Bank_Rec.Payee, 
              CASE WHEN CONVERT(INT, Ac_Bank_Rec.Debit) = 0 THEN 'Credit' ELSE 'Debit' END AS [Debit], 
              [Amount] = Ac_Bank_Rec.Amount 
              FROM Ac_Bank_Codes 
              INNER JOIN Ac_Bank_Rec ON Ac_Bank_Rec.Bank_Code=Ac_Bank_Codes.Code and Ac_Bank_Rec.Branch = Ac_Bank_Codes.Branch 
              INNER JOIN Branches ON Ac_Bank_Rec.Branch=Branches.Code 
              WHERE Ac_Bank_Rec.Client_Code = '{0}' AND Ac_Bank_Rec.Matter_No = {1} 
              ORDER BY Ac_Bank_Rec.ChequeDate""".format(lbl_EntRef.Content, lbl_MatNo.Content) 

  _tikitDbAccess.Open(mySQL)
  items = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          Bank = 'None' if dr.IsDBNull(0) else dr.GetString(0)
          TransNo = 'None' if dr.IsDBNull(1) else dr.GetString(1)
          BatchNo = 'None' if dr.IsDBNull(2) else dr.GetString(2)
          ChequeDate = 0 if dr.IsDBNull(3) else dr.GetValue(3)
          Ref = 'None' if dr.IsDBNull(4) else dr.GetString(4)
          Payee = 'None' if dr.IsDBNull(5) else dr.GetString(5)
          Debit = 'None' if dr.IsDBNull(6) else dr.GetValue(6)
          Amount = 0 if dr.IsDBNull(7) else dr.GetValue(7)
          items.append(UnclearedBankReceviedItemsObject(Bank, TransNo, BatchNo, ChequeDate, Ref, Payee, Debit, Amount))

  dr.Close()
  _tikitDbAccess.Close()      

  dg_UnclearedBankRecItems.ItemsSource = items
  UnclearedBankRecItemsLabel.Content = str(dg_UnclearedBankRecItems.Items.Count)
  lbl_UnclearedBankRecs.Content = str(dg_UnclearedBankRecItems.Items.Count)
  return True
# ----- UNCLEARED BANK RECEVIED ITEMS END ----

# ----- UNPROCESSED SLIPS START -----
class UnprocessedSlipsUpdateObject(object):
  def __init__(self, PostingType, Status, Originator, Ref, Narrative, Amount, ApprovalUser1, ApprovalUser2):
    self.PostingType = PostingType
    self.Status = Status
    self.Originator = Originator
    self.Ref = Ref
    self.Narrative = Narrative
    self.Amount = Amount
    self.ApprovalUser1 = ApprovalUser1
    self.ApprovalUser2 = ApprovalUser2
   
  def __getitem__(self, index):
    return None

def refresh_UnprocessedSlips():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False
    
  mySQL = """SELECT [Posting Type] = Ac_Posting_Slips.PostingCode, 
              CASE WHEN Ac_Posting_Slips.Status = 'U' THEN 'Unprocessed' WHEN Ac_Posting_Slips.Status = 'A' THEN 'Requires Approving' WHEN Ac_Posting_Slips.Status in ('P','D') THEN 'Remove from Housekeeping (Originators)' WHEN Ac_Posting_Slips.Status = 'C' THEN 'Remove from Housekeeping (Posted)'   END AS [Status], 
              [Originator] = Ac_Posting_Slips.Originator, 
              [Ref] = Ac_Posting_Slips.Ref, 
              [Narrative] = Ac_Posting_Slips.Narrative1 + Ac_Posting_Slips.Narrative2 + Ac_Posting_Slips.Narrative3, 
              [Amount] = Ac_Posting_Slips.TotalAmount, 
              [Approval User 1] = Ac_Posting_Slips.ApprovalUser, 
              [Approval User 2] = Ac_Posting_Slips.ApprovalUser_2 
              FROM Ac_Posting_Slips 
              WHERE Ac_Posting_Slips.Client1Code ='{0}' AND Ac_Posting_Slips.Client1MatterNo = {1} 
              ORDER BY Ac_Posting_Slips.PostingCode """.format(lbl_EntRef.Content, lbl_MatNo.Content)
  
  _tikitDbAccess.Open(mySQL)
  items = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          PostingType = 'None' if dr.IsDBNull(0) else dr.GetString(0)
          Status = 'None' if dr.IsDBNull(1) else dr.GetString(1)
          Originator = 'None' if dr.IsDBNull(2) else dr.GetString(2)
          Ref = 'None' if dr.IsDBNull(3) else dr.GetString(3)
          Narrative = 'None' if dr.IsDBNull(4) else dr.GetString(4)
          Amount = 'None' if dr.IsDBNull(5) else dr.GetValue(5)
          ApprovalUser1 = 'None' if dr.IsDBNull(6) else dr.GetString(6)
          ApprovalUser2 = 'None' if dr.IsDBNull(7) else dr.GetString(7)
          items.append(UnprocessedSlipsUpdateObject(PostingType, Status, Originator, Ref, Narrative, Amount, ApprovalUser1, ApprovalUser2))
   
  dr.Close()
  _tikitDbAccess.Close()

  dg_UnprocessedSlips.ItemsSource = items
  UnprocessedSlipsLabel.Content = str(dg_UnprocessedSlips.Items.Count)
  lbl_UnprocessedSlips.Content = str(dg_UnprocessedSlips.Items.Count)
  return True
# ----- UNPROCESSED SLIPS END -----

# ----- NON ZERO BALANCES START -----
class NonzeroBalancesUpdateObject(object):
  def __init__(self, BalanceType, BalanceValue, TextColour):
    self.BalanceType = BalanceType
    self.BalanceValue = BalanceValue
    self.TextColour = TextColour
    return
   
    #ClientBalance, OfficeBalance, UnpaidBillBalance, UnpaidDisbBalance, MemoBalance, AnticipatedBalance, NYPBalance, DepositBalance, UnbilledTimeBalance, UnbilledTimeBalanceValue):
    #self.ClientBalance = ClientBalance
    #self.OfficeBalance = OfficeBalance
    #self.UnpaidBillBalance = UnpaidBillBalance
    #self.UnpaidDisbBalance = UnpaidDisbBalance
    #self.MemoBalance = MemoBalance
    #self.AnticipatedBalance = AnticipatedBalance
    #self.NYPBalance = NYPBalance
    #self.DepositedBalance = DepositBalance
    #self.UnbilledTimeBalance = UnbilledTimeBalance
    #self.UnbilledTimeBalanceValue = UnbilledTimeBalanceValue
   
  def __getitem__(self, index):
    return None

def refresh_NonzeroBalances():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False

  mySQL = """SELECT 
                v.BalanceType,
                v.BalanceValue,
                v.TextColour
            FROM Matters
            CROSS APPLY (VALUES 
                ('Client Balance', Client_Ac_Balance, CASE WHEN Client_Ac_Balance > 0 THEN 'Red' ELSE 'Black' END),
                ('Office Balance', Office_Ac_Balance, CASE WHEN Office_Ac_Balance > 0 THEN 'Red' ELSE 'Black' END),
                ('Unpaid Bill Balance', UnpaidBillBalance, CASE WHEN UnpaidBillBalance > 0 THEN 'Red' ELSE 'Black' END),
                ('Unpaid Disb Balance', UnbilledDisbBalance, CASE WHEN UnbilledDisbBalance > 0 THEN 'Red' ELSE 'Black' END),
                ('Memo Balance', Memo_Balance, CASE WHEN Memo_Balance > 0 THEN 'Red' ELSE 'Black' END),
                ('Anticipated Balance', AnticipatedDisbsBalance, CASE WHEN AnticipatedDisbsBalance > 0 THEN 'Red' ELSE 'Black' END),
                ('NYP Balance', NYP_Balance, CASE WHEN NYP_Balance > 0 THEN 'Red' ELSE 'Black' END),
                ('Deposit Balance', Depost_Ac_Balance, CASE WHEN Depost_Ac_Balance > 0 THEN 'Red' ELSE 'Black' END),
                ('Unbilled Time Balance', UnbilledTimeBalance, CASE WHEN UnbilledTimeBalance > 0 THEN 'Red' ELSE 'Black' END),
                ('Unbilled Time Value', UnbilledTimeBalanceValue, CASE WHEN UnbilledTimeBalanceValue > 0 THEN 'Red' ELSE 'Black' END)
            ) v(BalanceType, BalanceValue, TextColour)
            WHERE EntityRef = '{0}' AND Number = {1};""".format(lbl_EntRef.Content, lbl_MatNo.Content)

  #mySQL = """SELECT [Client Balance] = Client_Ac_Balance, 
  #            [Office Balance] = Office_Ac_Balance, 
  #            [Unpaid Bill Balance] = UnpaidBillBalance, 
  #            [Unpaid Disb Balance] = UnbilledDisbBalance, 
  #            [Memo Balance] = Memo_Balance, 
  #            [Anticipated Balance] = AnticipatedDisbsBalance, 
  #            [NYP Balance] = NYP_Balance, 
  #            [Deposit Balance] = Depost_Ac_Balance, 
  #            [Unbilled Time Balance] = UnbilledTimeBalance, 
  #            [Unbilled Time Value] = UnbilledTimeBalanceValue 
  #            FROM Matters 
  #            WHERE EntityRef = '{0}' AND Number = {1}""".format(lbl_EntRef.Content, lbl_MatNo.Content)

  _tikitDbAccess.Open(mySQL)
  items = []
  countNonZero = 0
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          tBalType = '' if dr.IsDBNull(0) else dr.GetString(0)
          tBalValue = 0 if dr.IsDBNull(1) else dr.GetValue(1)
          tTextColour = '' if dr.IsDBNull(2) else dr.GetString(2)

          items.append(NonzeroBalancesUpdateObject(BalanceType=tBalType, BalanceValue=tBalValue, TextColour=tTextColour))

          # increment count of only non-zero values
          countNonZero += 1 if tBalValue > 0 else 0
          # TODO: just had a thought that now we're switching to ROWS instead of COLUMNS, it may make sense to EXCLUDE those = 0 from the get go
          # TODO   so we needn't worry about cell colouring on DataGrid, as only those items shown are the pertinent ones
          # TODO   may want to check with a Fee Earner though, as it may be that they DO want to still see that other values ARE zero?

          #ClientBalance = 'None' if dr.IsDBNull(0) else dr.GetValue(0)
          #OfficeBalance = 'None' if dr.IsDBNull(1) else dr.GetValue(1)
          #UnpaidBillBalance = 'None' if dr.IsDBNull(2) else dr.GetValue(2)
          #UnpaidDisbBalance = 'None' if dr.IsDBNull(3) else dr.GetValue(3)
          #MemoBalance = 'None' if dr.IsDBNull(4) else dr.GetValue(4)
          #AnticipatedBalance = 'None' if dr.IsDBNull(5) else dr.GetValue(5)
          #NYPBalance = 'None' if dr.IsDBNull(6) else dr.GetValue(6)
          #DepositBalance = 'None' if dr.IsDBNull(7) else dr.GetValue(7)
          #UnbilledTimeBalance = 'None' if dr.IsDBNull(8) else dr.GetValue(8)
          #UnbilledTimeBalanceValue = 'None' if dr.IsDBNull(9) else dr.GetValue(9)
          #
          ## Was having an issue where it would return a row of all 0
          ## The code below goes over the values and checks if theres a value thats not equal to 0, if so then
          ## c is incremented and the row is added to the datagrid
          #c = 0
          #tmp = [ClientBalance, OfficeBalance, UnpaidBillBalance, UnpaidDisbBalance, MemoBalance, AnticipatedBalance, NYPBalance, DepositBalance, UnbilledTimeBalance, UnbilledTimeBalance]
          # 
          #for val in tmp:
          #  if int(val) != 0:
          #    c +=1
          #
          #if c > 0:
          #  items.append(NonzeroBalancesUpdateObject(ClientBalance, OfficeBalance, UnpaidBillBalance, UnpaidDisbBalance, MemoBalance, AnticipatedBalance, NYPBalance, DepositBalance, UnbilledTimeBalance, UnbilledTimeBalanceValue))

    dr.Close()
  _tikitDbAccess.Close()

  dg_NonZeroBalances.ItemsSource = items
  NonZeroBalancesLabel.Content = str(countNonZero)
  lbl_NonZeroBals.Content = str(countNonZero)
  return True
# ----- NON ZERO BALANCES END -----

  
# ----- POST TO REVIEW START -----
class PostToReviewObject(object):
  def __init__(self, User, DateAdded, CaseId, CaseDesc, CaseAgenda):
    self.User = User
    self.DateAdded = DateAdded
    self.CaseId = CaseId
    self.CaseDesc = CaseDesc
    self.CaseAgenda = CaseAgenda
   
  def __getitem__(self, index):
    return None

def refresh_PostToReview():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False
  
  mySQL = """SELECT [User] = pr.[FeeEarnerRef], 
              [Date Added] = pr.[AddedDate], 
              [Case ID] = pr.StepID, 
              [Case Description] = ci.Description, 
              [Case Agenda] = ag.Description 
              FROM PostToReview pr 
              LEFT OUTER JOIN Cm_CaseItems ci ON ci.ItemID=pr.StepID 
              LEFT OUTER JOIN Cm_CaseItems ag ON ci.ParentID=ag.ItemID 
              LEFT OUTER JOIN Cm_Agendas ags ON ag.ItemID=ags.ItemID 
               WHERE ags.EntityRef = '{0}' AND ags.MatterNo = {1} AND pr.ReadDate IS NULL """.format(lbl_EntRef.Content, lbl_MatNo.Content)

  _tikitDbAccess.Open(mySQL)
  items = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          User = 'None' if dr.IsDBNull(0) else dr.GetString(0)
          DateAdded = 'None' if dr.IsDBNull(1) else dr.GetValue(1)
          CaseId = 'None' if dr.IsDBNull(2) else dr.GetValue(2)
          CaseDesc = 'None' if dr.IsDBNull(3) else dr.GetString(3)
          CaseAgenda = 'None' if dr.IsDBNull(4) else dr.GetString(4)

          items.append(PostToReviewObject(User, DateAdded, CaseId, CaseDesc, CaseAgenda))
     
  dr.Close()
  _tikitDbAccess.Close()

  dg_PostToReview.ItemsSource = items
  PostToReviewLabel.Content = str(dg_PostToReview.Items.Count)
  lbl_PostToReview.Content = str(dg_PostToReview.Items.Count)
  return True
# ----- POST TO REVIEW END -----

# ----- CHECKED OUT DOCUMENTS START ----
class CheckedOutDocumentsObject(object):
  def __init__(self, CaseStepId, CaseDescription, CheckedOutBy, CaseAgenda):
    self.CaseStepId = CaseStepId
    self.CaseDescription = CaseDescription
    self.CheckedOutBy = CheckedOutBy
    self.CaseAgenda = CaseAgenda
   
  def __getitem__(self, index):
    if index == 'ID':
      return self.CaseStepId

def refresh_CheckedOutDocuments():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False

  mySQL = """SELECT [Case Step ID] = cm.[ItemID], 
              [Case Description] = ci.[Description], 
              [Checked Out By] = cm.[CVSLockUser], 
              [Case Agenda] = ag.Description 
              FROM CM_Steps cm 
              LEFT OUTER JOIN Cm_CaseItems ci ON ci.ItemID=cm.ItemID 
              LEFT OUTER JOIN Cm_CaseItems ag ON ci.ParentID=ag.ItemID 
              LEFT OUTER JOIN Cm_Agendas ags ON ag.ItemID=ags.ItemID 
              WHERE ags.EntityRef = '{0}' AND ags.MatterNo = {1} AND cm.CVSLockUser <> '' """.format(lbl_EntRef.Content, lbl_MatNo.Content)

  _tikitDbAccess.Open(mySQL)
  items = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          CaseStepId = 'None' if dr.IsDBNull(0) else dr.GetValue(0)
          CaseDescription = 'None' if dr.IsDBNull(1) else dr.GetString(1)
          CheckedOutBy = 'None' if dr.IsDBNull(2) else dr.GetString(2)
          CaseAgenda = 'None' if dr.IsDBNull(3) else dr.GetString(3)

          items.append(CheckedOutDocumentsObject(CaseStepId, CaseDescription, CheckedOutBy, CaseAgenda))
      
  dr.Close()
  _tikitDbAccess.Close()

  dg_CheckedOutDocuments.ItemsSource = items
  CheckedOutDocumentsLabel.Content = str(dg_CheckedOutDocuments.Items.Count)
  lbl_CheckedOutDocs.Content = str(dg_CheckedOutDocuments.Items.Count)
  return True
# ----- CHECKED OUT DOCUMENTS END -----

# ----- UNDERTAKINGS LIST -----
class Undertakings(object):
  def __init__(self, myID, myStatus, myOrigUser, myRespFE, myAmyPay, myAmtRec, myType, myStageDue, myMadeDate, myDesc, myPurpose, myDischDate, myDischNote):
    self.uID = myID
    self.uStatus = myStatus
    self.uOrigUser = myOrigUser
    self.uRespFeeEarner = myRespFE
    self.uAmountPayable = myAmyPay
    self.uAmountRec = myAmtRec
    self.uType = myType
    self.uStageDue = myStageDue
    self.uMadeDate = myMadeDate
    self.uDesc = myDesc
    self.uPurpose = myPurpose
    self.uDischargeDate = myDischDate
    self.uDischargeNotes = myDischNote
    
  
  def __getitem__(self, index):
    
    if index == 'ID':
      return self.uID
    elif index == 'Status':
      return self.uStatus
    elif index == 'Desc':
      return self.uDesc
      
      
def refresh_UndertakingsList():

  # if no entity ref or matter no, then return False
  if lbl_EntRef.Content == '' or lbl_MatNo.Content == '':
    return False
  
  mySQL = """SELECT [ID] = UR.ID, [Status] = CASE WHEN UR.Status = 1 THEN 'Active' 
                                                  WHEN UR.Status = 2 THEN 'Discharged' 
                                                  WHEN UR.Status = 3 THEN 'Received' 
                                                  WHEN UR.Status = 4 THEN 'Cancelled' 
                                                  ELSE '' END, 
              [Originating User] = Users.FullName, 
              [Responsible Fee Earner] = Users_1.FullName, 
              [Amount Payable] = UR.AmountPayable, 
              [Amount Receivable] = UR.AmountReceivable, 
              [Type] = UD.Description, 
              [Stage Due] = UD_1.Description, 
              [Undertaking Made Date] = UR.UndertakingMadeDate, 
              [Description] = UR.UndertakingDescription, 
              [Purpose] = UR.Purpose, 
              [Discharge Date] = UR.DischargeDate, 
              [Discharge Notes] = UR.DischargeNotes 
              FROM UndertakingsRegister UR 
              LEFT JOIN Users ON UR.OriginatingUser = Users.Code 
              LEFT JOIN UndertakingsDescriptions UD ON UR.TypeRef = UD.ID 
              LEFT JOIN Users Users_1 ON UR.ResponsibleFERef = Users_1.Code 
              INNER JOIN UndertakingsDescriptions UD_1 ON UR.StageDueRef = UD_1.ID 
              WHERE UR.EntityRef = '{0}' AND UR.MatterNo = {1} AND UR.Status = 1 
              ORDER BY UR.UndertakingMadeDate """.format(lbl_EntRef.Content, lbl_MatNo.Content)

  #MessageBox.Show('mySQL: ' + mySQL)
  _tikitDbAccess.Open(mySQL)
  uItems = []
  
  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        if not dr.IsDBNull(0):
          aID = 0 if dr.IsDBNull(0) else dr.GetValue(0)
          aStatus = '' if dr.IsDBNull(1) else dr.GetString(1)
          aOrigUser = '' if dr.IsDBNull(2) else dr.GetString(2)
          aRespFE = '' if dr.IsDBNull(3) else dr.GetString(3)
          aAmtPay = 0 if dr.IsDBNull(4) else dr.GetValue(4)
          aAmtRec = 0 if dr.IsDBNull(5) else dr.GetValue(5)
          aType = '' if dr.IsDBNull(6) else dr.GetString(6)
          aStageDue = '' if dr.IsDBNull(7) else dr.GetString(7)
          aMadeDate = '' if dr.IsDBNull(8) else dr.GetValue(8)
          aDesc = '' if dr.IsDBNull(9) else dr.GetString(9)
          aPurpose = '' if dr.IsDBNull(10) else dr.GetString(10)
          aDischDate = '' if dr.IsDBNull(11) else dr.GetValue(11)
          aDischNote = '' if dr.IsDBNull(12) else dr.GetString(12)
          
          uItems.append(Undertakings(aID, aStatus, aOrigUser, aRespFE, aAmtPay, aAmtRec, aType, aStageDue, aMadeDate, aDesc, aPurpose, aDischDate, aDischNote))
    
  dr.Close()
  _tikitDbAccess.Close()

  dg_Undertakings.ItemsSource = uItems
  UndertakingsLabel.Content = str(dg_Undertakings.Items.Count)
  lbl_Undertakings.Content = str(dg_Undertakings.Items.Count)
  return True

# ----- UNDERTAKINGS LIST END -----


def Delete_Tasks(s, event):

  tmpEntRef = lbl_EntRef.Content 
  tmpMatNo = lbl_MatNo.Content 

  if dg_OutstandingTasks.Items.Count > 0:
    msg = "Are you sure you want to delete all Outstanding Tasks?"
    myResult = MessageBox.Show(msg, 'Delete Tasks?', MessageBoxButtons.YesNo)
    
    if myResult == DialogResult.No:
      return False
      
    # Firstly, get count of items in each table...
    count_Steps = _tikitResolver.Resolve("[SQL: SELECT COUNT(ItemID) FROM Cm_Steps WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_Steps_AH = _tikitResolver.Resolve("[SQL: SELECT COUNT(ItemID) FROM Cm_Steps_ActionHistory WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_StepGroupFolders = _tikitResolver.Resolve("[SQL: SELECT COUNT(StepID) FROM Cm_StepGroupFolders WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_StepLocks = _tikitResolver.Resolve("[SQL: SELECT COUNT(StepID) FROM Cm_Steps_Locks WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_TimePostings = _tikitResolver.Resolve("[SQL: SELECT COUNT(StepItemID) FROM Cm_TimePostings WHERE StepItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_CaseItems = _tikitResolver.Resolve("[SQL: SELECT COUNT(ItemID) FROM Cm_CaseItems WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
      
    # Delete Steps items (if count of items greater than zero)...
    if count_Steps > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_Steps WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_Steps_AH > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_Steps_ActionHistory WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_StepGroupFolders > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_StepGroupFolders WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_StepLocks > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_Steps_Locks WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_TimePostings > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_TimePostings WHERE StepItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_CaseItems > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_CaseItems WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Tasks UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupTasks) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
      
    # Now to UPDATE Diary_* tables to trigger (in-built) auto-delete mechanism (eg: so it deletes from Outlook too)
    _tikitResolver.Resolve("[SQL: UPDATE Diary_Tasks SET DeleteThis = 1 WHERE EntityRef = '{0}' AND MatterNoRef = {1}]".format(tmpEntRef, tmpMatNo))
    _tikitResolver.Resolve("[SQL: UPDATE Diary_GroupTasks SET DeleteThis = 1 WHERE EntityRef = '{0}' AND MatterNoRef = {1}]".format(tmpEntRef, tmpMatNo))
      
    # OLD CODE - CONSIDER DELETING
    #mySQL = "[SQL: DELETE FROM Diary_Tasks WHERE EntityRef = '" + _tikitEntity + "' AND MatterNoRef = " + str(_tikitMatter) + "]"
      
    #MessageBox.Show('Delete Tasks SQL: ' + mySQL)
    #_tikitResolver.Resolve(mySQL)
      
    #Group Tasks
    #mySQL = "[SQL: DELETE FROM Diary_GroupTasks WHERE EntityRef = '" + _tikitEntity + "' AND MatterNoRef = " + str(_tikitMatter) + "]"
    #MessageBox.Show('Delete GROUP Tasks SQL: ' + mySQL)
    #_tikitResolver.Resolve(mySQL)
      
  else:
    MessageBox.Show("No Tasks outstanding to delete!")
  return True    


def Delete_Appointments(s, event):

  tmpEntRef = lbl_EntRef.Content 
  tmpMatNo = lbl_MatNo.Content 

  if dg_OutstandingAppointments.Items.Count > 0:
    msg = "Are you sure you want to delete all Outstanding Appointments?"
    myResult = MessageBox.Show(msg, 'Delete Appointments?', MessageBoxButtons.YesNo)
    
    if myResult == DialogResult.No:
      return False

    # Firstly, get count of items in each table...
    count_Steps = _tikitResolver.Resolve("[SQL: SELECT COUNT(ItemID) FROM Cm_Steps WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_Steps_AH = _tikitResolver.Resolve("[SQL: SELECT COUNT(ItemID) FROM Cm_Steps_ActionHistory WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_StepGroupFolders = _tikitResolver.Resolve("[SQL: SELECT COUNT(StepID) FROM Cm_StepGroupFolders WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_StepLocks = _tikitResolver.Resolve("[SQL: SELECT COUNT(StepID) FROM Cm_Steps_Locks WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_TimePostings = _tikitResolver.Resolve("[SQL: SELECT COUNT(StepItemID) FROM Cm_TimePostings WHERE StepItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    count_CaseItems = _tikitResolver.Resolve("[SQL: SELECT COUNT(ItemID) FROM Cm_CaseItems WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
      
    # Delete Steps items (if count of items greater than zero)...
    if count_Steps > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_Steps WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_Steps_AH > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_Steps_ActionHistory WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_StepGroupFolders > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_StepGroupFolders WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_StepLocks > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_Steps_Locks WHERE StepID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_TimePostings > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_TimePostings WHERE StepItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
    if count_CaseItems > 0:
      _tikitResolver.Resolve("[SQL: DELETE FROM Cm_CaseItems WHERE ItemID IN (SELECT CaseItemRef FROM (SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_Appointments UNION SELECT CaseItemRef, EntityRef, MatterNoRef FROM Diary_GroupAppointments) as myTabke WHERE EntityRef = '{0}' AND MatterNoRef = {1})]".format(tmpEntRef, tmpMatNo))
      
    # Now to UPDATE Diary_* tables to trigger (in-built) auto-delete mechanism (eg: so it deletes from Outlook too)
    _tikitResolver.Resolve("[SQL: UPDATE Diary_Appointments SET DeleteThis = 1 WHERE EntityRef = '{0}' AND MatterNoRef = {1}]".format(tmpEntRef, tmpMatNo))
    _tikitResolver.Resolve("[SQL: UPDATE Diary_GroupAppointments SET DeleteThis = 1 WHERE EntityRef = '{0}' AND MatterNoRef = {1}]".format(tmpEntRef, tmpMatNo))
      
    # OLD CODE - CONSIDER DELETING
    #mySQL = "[SQL: DELETE FROM Diary_Appointments WHERE EntityRef = '" + _tikitEntity + "' AND MatterNoRef = " + str(_tikitMatter) + "]"
    #MessageBox.Show('Delete Appointments SQL: ' + mySQL)
    #_tikitResolver.Resolve(mySQL)
      
    # Group Appointments
    #mySQL = "[SQL: DELETE FROM Diary_GroupAppointments WHERE EntityRef = '" + _tikitEntity + "' AND MatterNoRef = " + str(_tikitMatter) + "]"
    #MessageBox.Show('Delete GROUP Appointments SQL: ' + mySQL)
    #_tikitResolver.Resolve(mySQL)
    
  else:
    MessageBox.Show("No Appointments oustanding to delete!")
  return True


def Discharge_Undertakings(s, event):

  tmpEntRef = lbl_EntRef.Content 
  tmpMatNo = lbl_MatNo.Content 

  if dg_Undertakings.SelectedIndex == -1:
    dg_Undertakings.SelectedIndex = 0
  
  itemID = dg_Undertakings.SelectedItem['ID']
  tmpItem = "({0}): {1}".format(itemID, dg_Undertakings.SelectedItem['Desc'])
  
  msg = "Are you sure you want mark the selected Undertaking as Discharged?\n" + tmpItem
  myResult = MessageBox.Show(msg, 'Mark Undertaking as Discharged?', MessageBoxButtons.YesNo)
  
  if myResult == DialogResult.Yes:
    FERef = _tikitResolver.Resolve("[SQL: SELECT FeeEarnerRef FROM Matters WHERE EntityRef = '{0}' AND Number = {1}]".format(tmpEntRef, tmpMatNo))
  
    mySQL = """[SQL: UPDATE UndertakingsRegister SET Status = 2, 
                DischargeDate = GETDATE(), 
                DischargeFERef = '{0}', 
                DischargeNotes = DischargeNotes + '*End of Workflow Step*' 
                WHERE EntityRef = '{1}' AND MatterNo = {2}
                 AND ID = {3}]""".format(FERef, tmpEntRef, tmpMatNo, itemID)
    
    #MessageBox.Show('Discharge Undertaking SQL: ' + mySQL)
  
    _tikitResolver.Resolve(mySQL)
    refresh_UndertakingsList(s, event)
    return True
  else:
    return False


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

  # else if already a string, assign passed date directly into newDate 
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
        #MessageBox.Show('Original: ' + str(varDate) + '\nFinal: ' + testStr)
        #newDate1 = datetime.datetime(int(tmpYear), int(tmpMonth), int(tmpDay))
        #finalStr = newDate1.strftime("%Y-%m-%d")
      finalStr = testStr

    return finalStr


def Clear_CheckedOut_Documents(s, event):

  if dg_CheckedOutDocuments.Items.Count > 0:
    
    msg = "Are you sure you want to clear the 'checked-out' status of all documents?"
    myResult = MessageBox.Show(msg, 'Clear checked-out documents?', MessageBoxButtons.YesNo)
  
    if myResult == DialogResult.Yes:    
      for x in dg_CheckedOutDocuments.Items:
        tmpID = x.CaseStepId
        tmpSQL = "[SQL: UPDATE Cm_Steps SET CVSLockUser = '', ReadLock = 0, WriteLock = 0 WHERE ItemID = {0}]".format(tmpID)
        #MessageBox.Show('ID of item: ' + tmpSQL)
        _tikitResolver.Resolve(tmpSQL)
  return True

]]>
    </Init>
    <Loaded>
      <![CDATA[
#Define controls that will be used in all of the code

# Along top of XAML - listed controls from left-to-right
btn_Help = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Help')
btn_Help.Click += btn_Help_Clicked
btn_Refresh = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_Refresh')
btn_Refresh.Click += refreshWIPReviewDataGrid

cbo_FeeEarner = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_FeeEarner')
cbo_FeeEarner.SelectionChanged += cbo_FeeEarner_SelectionChanged
cbo_SortBy = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_SortBy')
cbo_SortBy.SelectionChanged += refreshWIPReviewDataGrid

#chk_ViewDetails = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_ViewDetails')
#chk_ViewDetails.Click += update_Details_Datagrids
btn_ViewArchiveDetails = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_ViewArchiveDetails')
btn_ViewArchiveDetails.Click += toggle_ViewArchiveMatterDetails

# Bulk update area
lbl_LastSubmittedDate = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_LastSubmittedDate')
# following is the textblock residing inside the 'Submit' button as we need text to wrap, and to change according to users 'canApproveOwn'

# Main DataGrid
dg_WIPReview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_WIPReview')
dg_WIPReview.SelectionChanged += update_Details_Datagrids
dg_WIPReview.CellEditEnding += cellEdit_Finished

# Controls to temporarily hold selected row data (so we can compare values when datagrid is updated)
lbl_tmpNote = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_tmpNote')
lbl_tmpWOReq = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_tmpWOReq')
lbl_tmpWOtype = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_tmpWOtype')

## DETAIL TABS BELOW ##
# following is actually the 'Details' section at the bottom of the XAML - used to 'show/hide' area as a whole
# ultimately - not sure why I put into a GroupBox control... StackPanel would look nicer without border (but job for another day)
grp_WIPReview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'grp_WIPReview')

# WIP (Unbilled Time) tab
ti_UnbilledTime = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_UnbilledTime')
dg_UnbilledTime = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_UnbilledTime')
dg_TimeUsers = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_TimeUsers')
dg_TimeUsers.SelectionChanged += refresh_Matter_UnbilledTime

# Unbilled Disbursements tab
ti_UnbilledDisbs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_UnbilledDisbs')
dg_UnbilledDisbs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_UnbilledDisbs')
lbl_NO_DISBS = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_NO_DISBS')

# Unpaid Bills tab
ti_UnpaidBills = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_UnpaidBills')
dg_UnpaidBills = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_UnpaidBills')
lbl_NO_BILLS = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_NO_BILLS')

# Client Ledger tab
ti_ClientLedger = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_ClientLedger')
dgClientLedger = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dgClientLedger')

# Case Docs tab (note - disabled section as out of memory issue)
ti_CaseDocs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_CaseDocs')
dg_CaseManagerDocs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_CaseManagerDocs')
dg_CaseManagerDocs.SelectionChanged += CaseDoc_SelectionChanged
#dg_CaseManagerDocs.CellDoubleClick += open_Selected_CaseDoc
btn_OpenCaseDoc = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_OpenCaseDoc')
btn_OpenCaseDoc.Click += open_Selected_CaseDoc
cbo_AgendaName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'cbo_AgendaName')
cbo_AgendaName.SelectionChanged += refresh_CaseDocs
btn_CaseDocsRefresh = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_CaseDocsRefresh')
btn_CaseDocsRefresh.Click += caseDocsPanel_refresh
lbl_OurRefCD = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_OurRefCD')


# New 'Archive Details for Selected Matter' tab
lbl_OurRef = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_OurRef')
lbl_ClientName = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ClientName')
lbl_MatDesc = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MatDesc')
lbl_EntRef = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_EntRef')
lbl_MatNo = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_MatNo')

# totals labels for the count of items in associated DataGrid
lbl_ChecklistCount = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ChecklistCount')
lbl_OSapptTask = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_OSapptTask')
lbl_Undertakings = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_Undertakings')
lbl_ForwardPostedItems = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_ForwardPostedItems')
lbl_UnclearedBankRecs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_UnclearedBankRecs')
lbl_UnprocessedSlips = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_UnprocessedSlips')
lbl_NonZeroBals = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_NonZeroBals')
lbl_PostToReview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_PostToReview')
lbl_CheckedOutDocs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'lbl_CheckedOutDocs')

ti_OsAppts = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_OsAppts')
dg_OutstandingAppointments = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_OutstandingAppointments')
OutstandingAppointmentsLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'OutstandingAppointmentsLabel')
btn_DeleteAppointments = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteAppointments')
btn_DeleteAppointments.Click += Delete_Appointments

ti_OsTasks = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_OsTasks')
dg_OutstandingTasks = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_OutstandingTasks')
OutstandingTasksLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'OutstandingTasksLabel')
btn_DeleteTasks = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DeleteTasks')
btn_DeleteTasks.Click += Delete_Tasks

ti_Undertaking = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_Undertaking')
dg_Undertakings = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_Undertakings')
UndertakingsLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'UndertakingsLabel')
btn_DischargeUndertakings = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_DischargeUndertakings')
btn_DischargeUndertakings.Click += Discharge_Undertakings

ti_ForPI = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_ForPI')
dg_ForwardPostedItems = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_ForwardPostedItems')
ForwardPostedItemsLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ForwardPostedItemsLabel')

ti_UnBankRec = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_UnBankRec')
dg_UnclearedBankRecItems = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_UnclearedBankRecItems')
UnclearedBankRecItemsLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'UnclearedBankRecItemsLabel')

ti_UnpSlip = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_UnpSlip')
dg_UnprocessedSlips = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_UnprocessedSlips')
UnprocessedSlipsLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'UnprocessedSlipsLabel')

ti_NZB = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_NZB')
dg_NonZeroBalances = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_NonZeroBalances')
NonZeroBalancesLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'NonZeroBalancesLabel')

ti_PtR = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_PtR')
dg_PostToReview = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_PostToReview')
PostToReviewLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'PostToReviewLabel')

ti_COD = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'ti_COD')
dg_CheckedOutDocuments = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'dg_CheckedOutDocuments')
CheckedOutDocumentsLabel = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'CheckedOutDocumentsLabel')
btn_ClearCheckedOutDocs = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'btn_ClearCheckedOutDocs')
btn_ClearCheckedOutDocs.Click += Clear_CheckedOut_Documents


# on load functions (moved into dedicated funtion)
myOnFormLoadEvent(_tikitSender, '')
]]>
    </Loaded>
  </fileclosure>
</tfb>