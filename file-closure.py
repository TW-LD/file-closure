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
from System.Windows.Data import Binding
from System.Windows.Forms import SelectionMode, MessageBox, MessageBoxButtons, DialogResult
from System.Windows.Input import KeyEventHandler
from System.Windows.Media import Brush, Brushes
from TWUtils import *

 # TODO: Refactor main WIPReview datagrid, to match expected fields for the file closure tool
 # TODO: draft tables and structure for the checklist.
# hello Bobba Fett

#Global Variables
UserIsHOD = False
all_ticked = False

## Main WIP Review DataGrid
class WIPreview(object):
    def __init__(self, myTicked, myRef, myClient, myMatDesc, myEntRef, myMatNo, myFENote, myTimeInactive, myLastUpdated):
        self.iTemTicked   = myTicked
        self.wOurRef      = myRef
        self.wClientName  = myClient
        self.wMatDesc     = myMatDesc
        self.wEntRef      = myEntRef
        self.wMatNo       = myMatNo
        self.wFENote      = myFENote
        self.wTimeInactive= myTimeInactive
        self.wLastUpdated = myLastUpdated
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
        return None


def refreshWIPReviewDataGrid(s, event):
  if cbo_FeeEarner.SelectedIndex == -1:
    MessageBox.Show("No items to show as a Fee Earner hasn't been selected from the drop-down.\n\nPlease note that only the Accounts department are able to select a different Fee Earner", "Refresh WIP Review List...")
    return

  # TODO: We need to redo the SQL below to grab the following fields: Our ref, Client Name, Matter Description, Time inactive, Archiving issues.

  # TODO: Need to add a table and SQL grab for fee earner notes for archiving.

  wip_SQL = """
  SELECT
      -- Existing columns
      '0-OurRef'        = E.ShortCode + '/' + CONVERT(VARCHAR, M.Number),
      '1-Client Name'   = E.Name,
      '2-Matter Desc'   = M.Description,
      '3-EntRef'        = M.EntityRef,
      '4-MatNo'         = M.Number,
      '5-ArchivingNote' = FCH.ArchivingNote,
      '6-TimeInactive' = 
       DATEDIFF(week,(
          SELECT MAX(d)
          FROM
          (
              -- 1) Last Bill Date
              SELECT M.LastBillPostingDate AS d

              UNION ALL

              -- 2) Last Time Posting
              SELECT M.LastTimePostingDate

              UNION ALL

              -- 3) Last Document (Case Manager) Date
              SELECT 
              (
                  SELECT TOP 1 CM.StepCreated
                  FROM View_CaseManagerMP CM
                  WHERE CM.EntityRef = M.EntityRef
                    AND CM.MatterRef = M.Number
                  ORDER BY CM.StepCreated DESC
              )
          ) AS AllDates
      ), GETDATE()),
      '7-LastUpdated' = FCH.LastUpdated

  FROM Matters M
      LEFT OUTER JOIN Entities E
          ON M.EntityRef = E.Code
      LEFT OUTER JOIN Usr_FileClosureHeader FCH
          ON FCH.EntityRef = M.EntityRef
          AND FCH.MatterNo = M.Number
  WHERE
      M.FeeEarnerRef = '{0}'
  ORDER BY [6-TimeInactive] DESC;
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

              wip_item = WIPreview(
                  iTicked,
                  iRef,
                  iClient,
                  iMatDesc,
                  iEntRef,
                  iMatNo,
                  iFENote,
                  iTimeInactive,
                  iLastUpdated
              )
              mItem.append(wip_item)
      else:
          mItem.append(
              WIPreview(False, "-N/A-", "-No Data-", "-No Data-", "", 0, 0, 0, "")
          )
      dr.Close()
  else:
      mItem.append(
          WIPreview(False, "-N/A-", "-No Data-", "-No Data-", "", 0, 0, 0, "")
      )

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
  MessageBox.Show("Error 2")
  if dg_WIPReview.SelectedIndex == -1:
    ti_UnbilledTime.Header = 'WIP (Unbilled Time) - No data'
    return

  tmpEntity = dg_WIPReview.SelectedItem['EntityRef']
  tmpMatter = dg_WIPReview.SelectedItem['MatterNo']
  if dg_TimeUsers.SelectedIndex == -1:
    dg_TimeUsers.SelectedIndex = 0
  showFor = dg_TimeUsers.SelectedItem['Code']
  tmpCountRows = 0

  time_SQL = """SELECT '0-Transaction Date' = TT.TransactionDate, '1-Quantity Of Time' = TT.QuantityOfTime / 60,
  '2-Value Of Time' = TT.ValueOfTime - TT.TimeValueBilled, '3-Activity Type' = ActT.Description,
  '4-Narrative' = TT.Narratives + TT.Narrative2 + TT.Narrative3,
  '5-Originator' = (SELECT FullName FROM Users U WHERE U.Code = TT.Originator)
  FROM TimeTransactions TT, ActivityTypes ActT
  WHERE TT.ActivityCodeRef = ActT.Code AND ActT.ChargeType = 'C' AND TT.TransactionType LIKE '%Time'
  AND TT.EntityRef = '{entity}' AND 
  TT.MatterNoRef = {matter} AND TT.ValueOfTime - TT.TimeValueBilled > 0 """.format(entity=tmpEntity, matter=tmpMatter)
  if str(showFor) != 'x':
    time_SQL += "AND TT.Originator = '" + str(showFor) + "' "
  time_SQL += "ORDER BY TT.TransactionDate DESC"

  # Open and store items in code
  _tikitDbAccess.Open(time_SQL)
  tItem = []

  if _tikitDbAccess._dr is not None:
    dr = _tikitDbAccess._dr
    if dr.HasRows:
      while dr.Read():
        iTDate = 0 if dr.IsDBNull(0) else dr.GetValue(0)  
        iQtyTime = 0 if dr.IsDBNull(1) else dr.GetDecimal(1)
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
  MessageBox.Show("Error 3")
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
  dNoData = False

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
    else:
      dItem.append(matterDisbs(0, "-N/A-", '', '-No Data for ' + dg_WIPReview.SelectedItem['OurRef'] + '-', 0, 0, 0, 0, 0))
      dNoData = True
    dr.Close()
  else:
    dItem.append(matterDisbs(0, "-N/A-", '', '-No Data-', 0, 0, 0, 0, 0))
    dNoData = True

  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_UnbilledDisbs.ItemsSource = dItem
  ti_UnbilledDisbs.Header = "Unbilled Disbursements ({0})".format(tmpCount)

  if dNoData == True:
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
  MessageBox.Show("Error 4")
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

  if dg_WIPReview.SelectedIndex == -1:
    lbl_tmpNote.Content = ''
    lbl_tmpWOReq.Content = ''
    lbl_tmpWOtype.Content = ''
    if chk_ViewDetails.IsChecked == True:
      grp_WIPReview.Header = 'Details for selected matter: (NO MATTER CURRENTLY SELECTED)'
  else:
    lbl_tmpNote.Content = dg_WIPReview.SelectedItem['FENote']
    if chk_ViewDetails.IsChecked == True:
      grp_WIPReview.Header = 'Details for selected matter: ' + dg_WIPReview.SelectedItem['OurRef'] + ' - ' + dg_WIPReview.SelectedItem['MatDesc'] + ' '
    dg_WIPReview.BeginEdit()

  if chk_ViewDetails.IsChecked == True:
    # update other datagrids and not fogetting to update the ticked status
    dg_WIPReview.Height = 431
    grp_WIPReview.Visibility = Visibility.Visible
    MessageBox.Show("Error 1")
    populate_TimeUsersToShow(s, event)
    refresh_Matter_UnbilledTime(s, event)
    refresh_Matter_UnbilledDisbs(s, event)
    refresh_Matter_UnpaidBills(s, event)
    refresh_dgClientLedger(s, event)
  else:
    grp_WIPReview.Visibility = Visibility.Hidden
    dg_WIPReview.Height = 790
    
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
  MessageBox.Show("Error 2")
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
        iToShow = 0 if dr.IsDBNull(0) else dr.GetString(0)
        iCode = 0 if dr.IsDBNull(1) else dr.GetString(1)
        iTotalTime = '' if dr.IsDBNull(2) else dr.GetValue(2)
        iValOfTime = '' if dr.IsDBNull(3) else dr.GetValue(3)  
        nTotalTime = getTextualTime(iTotalTime)

        nValOfTime = '{:,.2f}'.format(iValOfTime)
        
        myItem.append(timeUsers(iTick, iToShow, iCode, nTotalTime, nValOfTime))

      dr.Close()
  _tikitDbAccess.Close

  # Set 'Source' and close db connection
  dg_TimeUsers.ItemsSource = myItem
  return

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
  MessageBox.Show("Error 5")
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
    else:
      myItem.append(cls_dgClientLedger('', '', '', '', '', '', '', '', '', '', '', '', ''))
    dr.Close()
  _tikitDbAccess.Close

  dgClientLedger.ItemsSource = myItem
  if tmpCount == 0:
    ti_ClientLedger.Header = 'Client Ledger - No data'
  else:
    ti_ClientLedger.Header = 'Client Ledger - ' + str(tmpCount) + ' items'
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

  sSQL = "SELECT CI.ItemID, CI.Description, CI.CreationDate,  CMS.FileName, Agenda.Description "
  sSQL += "FROM Cm_CaseItems CI "
  sSQL += "INNER JOIN Cm_Steps CMS ON CMS.ItemID = CI.ItemID "
  sSQL += "INNER JOIN Cm_Agendas CMA ON CI.ParentID = CMA.ItemID "
  sSQL += "INNER JOIN Cm_CaseItems Agenda ON CMA.ItemID = Agenda.ItemID "
  sSQL += "WHERE CMS.FileName <> '' "

  if cbo_AgendaName.SelectedIndex == 0:
    # show ALL documents for current matter
    sSQL += "AND CMA.EntityRef = '" + str(dg_WIPReview.SelectedItem['EntityRef']) + "' AND CMA.MatterNo = " + str(dg_WIPReview.SelectedItem['MatterNo']) + " ORDER BY Agenda.ItemID, CI.ItemOrder "
    dg_CaseManagerDocs.Columns[0].Visibility = Visibility.Visible
  if cbo_AgendaName.SelectedIndex > 0:
    # just show docs for the selected Agenda
    tmpAgendaName = cbo_AgendaName.SelectedItem['ID']
    sSQL += "AND Agenda.ItemID = " + str(tmpAgendaName) + " ORDER BY CI.ItemOrder "
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

chk_ViewDetails = LogicalTreeHelper.FindLogicalNode(_tikitSender, 'chk_ViewDetails')
chk_ViewDetails.Click += update_Details_Datagrids

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

# on load functions (moved into dedicated funtion)
myOnFormLoadEvent(_tikitSender, '')
]]>
    </Loaded>
  </fileclosure>
</tfb>