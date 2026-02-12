Dim qtApp
Set qtApp = CreateObject("QuickTest.Application")

qtApp.Launch
qtApp.Visible = True

qtApp.Open "C:\Users\mridu\OneDrive\Documents\Functional Testing\GUITest1"

qtApp.Test.Run

qtApp.Test.Close
qtApp.Quit

Set qtApp = Nothing