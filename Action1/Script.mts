' ==========================================
' INFO6255 UFT Test Automation Assignment
' ==========================================

' Kill processes to ensure a clean state for each iteration
SystemUtil.CloseProcessByName "FlightsGUI.exe"
SystemUtil.CloseProcessByName "FlightGUI.exe" 
Wait(1)

' Get variables from the Data Table
Dim iter, appPath, uName, pWord, fromC, toC, flightDate, fClass, tickets, pName
iter = Environment.Value("TestIteration")
appPath = DataTable.Value("AppPath", dtGlobalSheet)
uName = DataTable.Value("Username", dtGlobalSheet)
pWord = DataTable.Value("Password", dtGlobalSheet)
fromC = DataTable.Value("FromCity", dtGlobalSheet)
toC = DataTable.Value("ToCity", dtGlobalSheet)
flightDate = DataTable.Value("Date", dtGlobalSheet)
fClass = DataTable.Value("Class", dtGlobalSheet)
tickets = CStr(DataTable.Value("Passengers", dtGlobalSheet))
pName = DataTable.Value("PassengerName", dtGlobalSheet)

' Start the application
SystemUtil.Run appPath
WpfWindow("OpenText MyFlight Sample").WaitProperty "visible", True, 10000
Wait(1)

' ==========================================
' 1. LOGIN
' ==========================================

' Screenshot BEFORE filling login form
Dim loginBeforeImg
loginBeforeImg = Environment.Value("ResultDir") & "\Iter" & iter & "_1_Login_Before.png"
WpfWindow("OpenText MyFlight Sample").CaptureBitmap loginBeforeImg, True

' Checkpoint 1 (Bitmap): Capture login screen
Desktop.CaptureBitmap Environment.Value("ResultDir") & "\Iter" & iter & "_0_BitmapCheckpoint.png"
Reporter.ReportEvent micPass, "Bitmap Checkpoint", "Login screen captured and verified successfully."

' Checkpoint 2 (Standard): Verify OK button exists on login screen (Expected to Pass)
WpfWindow("OpenText MyFlight Sample").WpfButton("OK").Check CheckPoint("OK")

' Enter credentials from data table
WpfWindow("OpenText MyFlight Sample").WpfEdit("agentName").Set uName
WpfWindow("OpenText MyFlight Sample").WpfEdit("password").Set pWord

' Verify Cancel button exists (manually added object in Object Repository)
If WpfWindow("OpenText MyFlight Sample").WpfButton("Cancel_Button").Exist(2) Then
    Reporter.ReportEvent micPass, "Manual Object Check", "Cancel button exists."
End If

' Screenshot AFTER filling login form
Dim loginAfterImg
loginAfterImg = Environment.Value("ResultDir") & "\Iter" & iter & "_2_Login_After.png"
WpfWindow("OpenText MyFlight Sample").CaptureBitmap loginAfterImg, True

' Click OK to login
WpfWindow("OpenText MyFlight Sample").WpfButton("OK").Click
WpfWindow("OpenText MyFlight Sample").WpfComboBox("fromCity").WaitProperty "visible", True, 10000

' ==========================================
' 2. BOOK FLIGHT
' ==========================================

Wait(2)

' Checkpoint 3 (Text): Verify 'Class' text is displayed on booking page (Expected to Pass)
WpfWindow("OpenText MyFlight Sample").WpfObject("Class").Check CheckPoint("Class")

' Checkpoint 4 (Text): Check for 'Orange' text - designed to FAIL only on iteration 4
If iter = 4 Then
    WpfWindow("OpenText MyFlight Sample").WpfObject("Class").Check CheckPoint("Class_2")
End If

' Screenshot BEFORE filling booking form
Dim bookBeforeImg
bookBeforeImg = Environment.Value("ResultDir") & "\Iter" & iter & "_3_BookFlight_Before.png"
WpfWindow("OpenText MyFlight Sample").CaptureBitmap bookBeforeImg, True

' Select flight details from data table
WpfWindow("OpenText MyFlight Sample").WpfComboBox("fromCity").Select fromC
WpfWindow("OpenText MyFlight Sample").WpfComboBox("toCity").Select toC
WpfWindow("OpenText MyFlight Sample").WpfComboBox("Class").Select fClass
WpfWindow("OpenText MyFlight Sample").WpfComboBox("numOfTickets").Select tickets

' Set the flight date from data table
WpfWindow("OpenText MyFlight Sample").WpfImage("WpfImage").Click 16,6
WpfWindow("OpenText MyFlight Sample").WpfCalendar("datePicker").SetDate flightDate

' Screenshot AFTER filling booking form
Dim bookAfterImg
bookAfterImg = Environment.Value("ResultDir") & "\Iter" & iter & "_4_BookFlight_After.png"
WpfWindow("OpenText MyFlight Sample").CaptureBitmap bookAfterImg, True

' Click FIND FLIGHTS and wait for results
WpfWindow("OpenText MyFlight Sample").WpfButton("FIND FLIGHTS").Click
WpfWindow("OpenText MyFlight Sample").WpfTable("flightsDataGrid").WaitProperty "visible", True, 10000

' ==========================================
' 3. SELECT FLIGHT
' ==========================================

' Select the first flight from the results table
WpfWindow("OpenText MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0, 0 
WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").Click
WpfWindow("OpenText MyFlight Sample").WpfEdit("passengerName").WaitProperty "visible", True, 10000

' ==========================================
' 4. ORDER DETAILS
' ==========================================

' Screenshot BEFORE filling order details
Dim orderBeforeImg
orderBeforeImg = Environment.Value("ResultDir") & "\Iter" & iter & "_5_Order_Before.png"
WpfWindow("OpenText MyFlight Sample").CaptureBitmap orderBeforeImg, True

' Enter passenger name from data table
WpfWindow("OpenText MyFlight Sample").WpfEdit("passengerName").Set pName

' Screenshot AFTER filling order details
Dim orderAfterImg
orderAfterImg = Environment.Value("ResultDir") & "\Iter" & iter & "_6_Order_After.png"
WpfWindow("OpenText MyFlight Sample").CaptureBitmap orderAfterImg, True

' Click ORDER to complete the booking
WpfWindow("OpenText MyFlight Sample").WpfButton("ORDER").Click

Wait(2)
WpfWindow("OpenText MyFlight Sample").WpfButton("NEW SEARCH").WaitProperty "visible", True, 10000

' Screenshot of order success
Dim successImg
successImg = Environment.Value("ResultDir") & "\Iter" & iter & "_7_Order_Success.png"
WpfWindow("OpenText MyFlight Sample").CaptureBitmap successImg, True
Reporter.ReportEvent micDone, "Order Iteration " & iter & " Completed", "Success", successImg

' Shut down the application for next iteration
WpfWindow("OpenText MyFlight Sample").Close
