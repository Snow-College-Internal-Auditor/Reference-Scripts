Dim listbox1$() As String
Dim list1$() As String

Begin Dialog DialogDemo 13,5,351,229,"Dialog Demo", .displayIt
  OKButton 22,180,40,14, "OK", .OKButton1
  CancelButton 75,180,40,14, "Cancel", .CancelButton1
  Text 10,10,30,10, "File:", .Text1
  Text 50,10,200,15, "Text", .txtFilename
  PushButton 260,12,10,10, "...", .btnFileSelect
  GroupBox 45,3,210,25, .GroupBox1
  Text 10,30,40,14, "High Value", .Text2
  TextBox 55,30,40,14, .txtHighValue
  Text 10,50,40,10, "Amount Field", .Text3
  DropListBox 55,49,215,11, listbox1$(), .drpAmountField
  GroupBox 7,70,300,32, "Select AP Tests", .GroupBox2
  CheckBox 14,80,40,14, "Select All", .chkAPSelAll
  CheckBox 70,80,40,14, "Test 1", .chkAPTest1
  CheckBox 131,80,40,14, "Test 2", .chkAPTest2
  CheckBox 189,80,40,14, "Test 3", .chkAPTest3
  GroupBox 7,107,300,29, "Select JV Tests", .GroupBox3
  CheckBox 14,117,40,14, "Select All", .chkJVSelAll
  CheckBox 70,117,40,14, "Test 1", .chkJVTest1
  CheckBox 131,117,40,14, "Test 2", .chkJVTest2
  CheckBox 189,117,40,14, "Test 3", .chkJVTest3
  GroupBox 7,141,300,29, "Amount Types", .GroupBox4
  OptionGroup .OptionButtonGroup1
  OptionButton 16,154,70,14, "Negative and Postive", .OptionButton1
  OptionButton 101,154,40,14, "Positve", .OptionButton2
  OptionButton 157,154,40,14, "Negative", .OptionButton3
  PushButton 135,180,40,14, "Help", .pbHelp
  Text 106,30,40,14, "Start Date:", .Text4
  TextBox 155,30,40,14, .txtStartDate
  PushButton 211,30,40,14, "Calendar", .btnCalendar
End Dialog

Begin Dialog HelpDialog 44,26,150,150,"Help Dialog", .helpMenuFunc
  Text 16,10,110,87, "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua. At vero eos et accusam et justo duo dolores et ea rebum. ", .Text1
  OKButton 19,109,40,14, "OK", .OKButton1
End Dialog

Begin Dialog dlgDatePicker 0,35,146,171,"Date Picker", .funCalendar
  PushButton 12,45,15,14, "31", .PB1
  PushButton 26,45,15,14, "31", .PB2
  PushButton 40,45,15,14, "31", .PB3
  PushButton 54,45,15,14, "31", .PB4
  PushButton 68,45,15,14, "31", .PB5
  PushButton 82,45,15,14, "31", .PB6
  PushButton 96,45,15,14, "31", .PB7
  PushButton 12,58,15,14, "31", .PB8
  PushButton 26,58,15,14, "31", .PB9
  PushButton 40,58,15,14, "31", .PB10
  PushButton 54,58,15,14, "31", .PB11
  PushButton 68,58,15,14, "31", .PB12
  PushButton 82,58,15,14, "31", .PB13
  PushButton 96,58,15,14, "31", .PB14
  PushButton 12,71,15,14, "31", .PB15
  PushButton 26,71,15,14, "31", .PB16
  PushButton 40,71,15,14, "31", .PB17
  PushButton 54,71,15,14, "31", .PB18
  PushButton 68,71,15,14, "31", .PB19
  PushButton 82,71,15,14, "31", .PB20
  PushButton 96,71,15,14, "31", .PB21
  PushButton 12,84,15,14, "31", .PB22
  PushButton 26,84,15,14, "31", .PB23
  PushButton 40,84,15,14, "31", .PB24
  PushButton 54,84,15,14, "31", .PB25
  PushButton 68,84,15,14, "31", .PB26
  PushButton 82,84,15,14, "31", .PB27
  PushButton 96,84,15,14, "31", .PB28
  PushButton 12,96,15,14, "31", .PB29
  PushButton 26,96,15,14, "31", .PB30
  PushButton 40,96,15,14, "31", .PB31
  PushButton 54,96,15,14, "31", .PB32
  PushButton 68,96,15,14, "31", .PB33
  PushButton 82,96,15,14, "31", .PB34
  PushButton 96,96,15,14, "31", .PB35
  PushButton 12,109,15,14, "31", .PB36
  PushButton 26,109,15,14, "31", .PB37
  PushButton 40,109,15,14, "31", .PB38
  PushButton 54,109,15,14, "31", .PB39
  PushButton 68,109,15,14, "31", .PB40
  PushButton 82,109,15,14, "31", .PB41
  PushButton 96,109,15,14, "31", .PB42
  PushButton 14,15,10,14, "<", .PBPrevious
  PushButton 101,14,10,14, ">", .PBNext
  Text 14,34,10,10, "Su", .Text1
  Text 29,34,10,10, "Mo", .Text1
  Text 42,34,10,10, "Tu", .Text1
  Text 56,34,10,10, "We", .Text1
  Text 70,34,10,10, "Th", .Text1
  Text 85,34,10,10, "Fr", .Text1
  Text 98,34,10,10, "Sa", .Text1
  CancelButton 17,131,40,14, "Cancel", .CancelButton1
  Text 41,18,50,9, "Text", .txtYearMonth
  DropListBox 69,2,26,11, list1$(), .lstYear
  Text 32,4,33,8, "Select Year", .Text3
  PushButton 68,131,40,14, "Current Date", .BTCurrentDate
End Dialog

















Option Explicit
Public Const LOCALE_SSHORTDATE = &H1F       '  short date format string
Public Const LOCALE_SDATE = &H1D            '  date separator
Public Const LOCALE_SYSTEM_DEFAULT& = &H800
Public Const LOCALE_USER_DEFAULT& = &H400

Dim sFilename As String
Dim bExitScript As Boolean
Dim bHelpDialogOpen As Boolean
Dim bCalendarOpen As Boolean 

'creating variables for input validation
Dim sHighValue As String
Dim sAmountField As String
Dim iAPTests(2) As Integer
Dim iJVTests(2) As Integer
Dim iAmountType As Integer

Dim sDayArray(41) As String 'string array to hold days of month
Dim sMonthArray(12) As String 'string array to hold months
Dim sYearArray(11) As String 'string ar
Dim iYear As Integer 'to track the year per the date picker
Dim iMonth As Integer  'to track the month per the date picker
Dim sDate As String 'to hold the date selected from teh date picker
Dim sDefaultDateFormat As String 'to hold the regional date format
Dim sDefaultDateSeperator 'to hold the date separator
Dim sDateDefault As String 'populated in getWeekday function, returns the date in the default format

Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Sub Main
	Call populateMonth() 'populate months
	Call menu()
	If Not bExitScript Then 
		If iAPTests(0) = 1 Then
			MsgBox "Test1"
		End If
		If iAPTests(1) = 1 Then
			MsgBox "Test2"
		End If
		If iAPTests(2) = 1 Then
			MsgBox "Test3"
		End If
	Else
		'MsgBox is how you can create a mesage box to diaplay to the user
		MsgBox "Script cancelled"
	End If
End Sub

Function menu()
	Dim dlg As DialogDemo
	Dim button As Integer 
	On Error Resume Next 
	sFilename = Client.CurrentDatabase.Name
	button = Dialog(dlg)
	If button = 0 Then bExitScript = True 
End Function

Function displayIt(ControllD$, Action%, SuppValue%)
	Dim bExitMenu As Boolean 
	Dim chk1 As Integer 
	Dim i As Integer 
	Dim chkjv1 As Integer
	Dim dlgHlp As HelpDialog
	'variable to reference calendar 
	Dim dlgDate As dlgDatePicker
	Dim button As Integer
	Dim bAPTests As Boolean
	Dim bJVTests As Boolean
	
	Select Case Action%
		Case 1
			If sFilename <>""Then
				Call getNumericFIeld
				DlgListBoxArray"drpAmountField", listbox1$()
			End If

		Case 2
			Select Case ControllD$
			
				Case "btnCalendar"
					If Not bCalendarOpen Then
						bCalendarOpen = True
						'this is grabing the information in the main dialog and sends it to the date picker
						sDate = DialogDemo.txtStartDate
						button = Dialog(dlgDate)
						bCalendarOpen = False
						'dlgText is a IDEA function to grab text from the user
						dlgText"txtStartDate",sDate
					End If
				Case "pbHelp"
					If Not bHelpDialogOpen Then
						bHelpDialogOpen = True
						button = Dialog(dlgHlp)
						bHelpDialogOpen = False
					End If		
				Case "chk1"
					If DlgValue("chk1")Then
						chk1 = 1
					Else 
						chk1 = 0
					End If
					For i = 1 To 3 
						DlgValue("ckTest" & i), chk1
					Next i
				Case "chkjv1"
					If DlgValue("chkjv1")Then
						chk1 = 1
					Else 
						chk1 = 0
					End If
					For i = 1 To 3 
						DlgValue("ckjvTest" & i), chk1
					Next i
				Case "btnFileSelect"
					sFilename = selectFile()
					If sFilename <>""Then
						Call getNumericFIeld
						DlgListBoxArray"drpAmountField", listbox1$()
					End If
				Case "OKButton1"
					For i = 0 To 2
						iAPTests(i) = DlgValue("ckTest"&(i + 1))
						If DlgValue("ckTest" & (i+1)) Then
							bAPTests = True
						End If
						iJVTests(i) = DlgValue("ckjvTest"&(i+1))
						If DlgValue("ckTest" & (i+1)) Then
							bJVTests = True
						End If
					Next i 
					
					sHighValue = DialogDemo.TextBox1
					'MsgBox "High Value" & dHighValue
					iAmountType = DialogDemo.OptionButtonGroup1
					'MsgBox "Amount Types" & iAmountType
					sAmountField = listbox1$(DialogDemo.drpAmountField + 1)
					'MsgBox "Amount Field" & sAmountField
					
					'cheking that these fields are not empty
					If sFilename = "" Then
						MsgBox "Please select a file"
					ElseIf sDate = "" Then
						MsgBox "Please enter a date"
					ElseIf Not bAPTests Then
						MsgBox"Please select at least one AP test"
					ElseIf Not bJVTests Then
						MsgBox"Please select at least one JV test"
					ElseIf Trim(sHighValue) = "" Then
						MsgBox"The high value field is empty, please enter an amount"
				 	ElseIf Not IsNumeric(sHighValue) Then
				 		MsgBox "Please enter a number in the high value field"
				 	'ElseIf Int(Val(sAmountField)) <> Val(sAmountField)Then
				 	Else 
				 		bExitMenu = True
					End If
					
					Client.OpenDatabase(sFilename)
				Case "CancelButton1"
					bExitMenu = True
					bExitScript = True
			End Select
	End Select
			
	If bExitMenu Then 
		displayIt = 0
	Else
		displayIt = 1
	End If
	
	If sFilename = ""Then
		DlgText "txtFilename","Please select a filename"
	Else
		DlgText "txtFilename",iSplit(sFilename,"","\",1,1)
	End If
	'this dispalys the dialog box for the calaneder so the user can not input into the date text box. 
	DlgEnable "txtStartDate",0
End Function

Function helpMenuFunc(ControllD$, Action%, SuppValue%)
	Dim msg As String 
	
	msg = "this is the help" & Chr(13) & Chr(10)
	msg = msg & "you can add any hlep in thsi section"& Chr(13) & Chr(10) & Chr(13) & Chr(10)
	msg = msg & "you can also use the  Chr(13) & Chr(10) to add carriage returns and have new lines"
	
	dlgText"Text1", msg
End Function

Function selectFile() As String
	Dim obj As Object 
	Set obj = Client.CommonDialogs
		selectFile = obj.FileExplorer()
	Set obj = Nothing 
End Function 

Function getNumericField()
	Dim db As database
	Dim table As table
	Dim field As field
	Dim i As Integer
	Dim bFirstTime As Boolean 
	
	Set db = Client.OpenDatabase(sFilename)
		Set table = db.TableDef
			bFirstTime = True 
			For i = 1 To table.count
				Set field = table.GetFieldAt(i)
				If field.isNumeric Then
					If bFirstTime Then
						bFirstTime = False
						ReDim listbox1$(1)
							listbox1$(1) = field.name
					Else
						ReDim preserve listbox1$(UBound(listbox1$)+1)
						listbox1$(UBound(listbox1$)) = field.name
					End If
				End If
			Next i
			Set field = Nothing
		Set table = Nothing
		db.close
	Set db = Nothing	
End Function 

Function funCalendar(ControlID$, Action%, SuppValue%)

	Dim iWeekDay As Integer 'holds the start day of the month
	Dim i, j As Integer
	Dim iNoOfDays As Integer 'numer of days in the month
	Dim bExitMenu As Boolean
	Dim bCurrentDate As Boolean 'flag to indicate if the script should use the current date
	Dim bUpdateYearList As Boolean 'flag to indicate if the year drop down should be updated as the year has changed
	
	bExitMenu  = FALSE
	bCurrentDate = FALSE
	
	Select Case action%
		'1 indicates that this is the first time the function is called
		Case 1 'check to see if there is a valid date and if so use it, if not use the current date
			If IsDate(sDate) Then
				iYear = Year(sDate)
				iMonth = Month(sDate)
				
				bUpdateYearList = TRUE

			Else
				bCurrentDate = TRUE
			End If
		Case 2
			Select Case ControlID$
				Case "CancelButton1"
					bExitMenu = TRUE
				Case "PBPrevious"
					'if the user clicks the previous arrow - < - then remove a month
					iYear = Year(DateSerial(iYear, iMonth  - 1, 1)) 
					iMonth = Month(DateSerial(iYear, iMonth - 1, 1))
				Case "PBNext"
					'if the user clicks the next arrow - > - then add a month
					iYear = Year(DateSerial(iYear, iMonth  + 1, 1))
					iMonth = Month(DateSerial(iYear, iMonth + 1, 1))
				Case "BTCurrentDate"
					'if the current date button is used default to the current date
					bCurrentDate = TRUE
				Case "lstYear"
					'if a year is selected from the drop down update the variables for the selected year and update the year dropdown
					iYear = Year(DateSerial(sYearArray(SuppValue%), iMonth, 1))
					iMonth = Month(DateSerial(sYearArray(SuppValue%), iMonth, 1))
					'update the year based on the new year
					bUpdateYearList = TRUE
				Case Else
					'if any of the date buttons are selected
					'if the button is blank do nothing
					If sDayArray(SuppValue% - 3) = "" Then 
						bExitMenu = FALSE
					Else
						'if a date is selected create the date and give it to the sDate variable which is a global variable and available from the main menu dialog
						sDate = iYear & "/"
						If iMonth < 10 Then
							sDate = sDate & "0" & iMonth & "/"
						Else
							sDate = sDate & iMonth & "/"
						End If
						If Len(sDayArray(SuppValue% - 3) ) = 1Then
							sDate = sDate & "0" & sDayArray(SuppValue% - 3)
						Else
							sDate = sDate & sDayArray(SuppValue% - 3)
						End If
						bExitMenu = TRUE
					End If
			End Select

			
	End Select
	
	If bExitMenu  Then
		funCalendar = 0
	Else
		funCalendar = 1
	End If
	
	'if the current date is selected then get the current year and month
	If bCurrentDate Then
		iYear = Year(Now())
		iMonth = Month(Now())
		bUpdateYearList = TRUE
	End If
	
	'update the year drop down based on the iYear variable, 5 years back and 5 years forward.
	If bUpdateYearList Then
		'populate the year event
		j = 0
		sYearArray(j) = "Year"
		j = j + 1
		For i = iYear - 5 To iYear 
			sYearArray(j)  = i
			j = j + 1
		Next i
		For i = iYear + 1 To iYear + 5 
			sYearArray(j)  = i
			j = j + 1
		Next i
		DlgListBoxArray "lstYear", sYearArray

	End If
	
	'clear the array
	For i = 0 To 41
		sDayArray(i) = "" 
	Next i
	
	'obtain the current week day 1 is Sunday 7 is Saturday
	iWeekDay = getWeekeday(iMonth, iYear)
	'obtain the number of days in the month
	iNoOfDays = dhDaysInMonth(sDateDefault )
	'populate the day array
	For i = 0 To iWeekDay - 1
		sDayArray(i) = "" 
	Next i
	For j = 1 To iNoOfDays
		sDayArray(iWeekDay -2 + j) = j
		i = i + 1
	Next j
	'put the year and month in the heading
	DlgText "txtYearMonth", sMonthArray(iMonth) & " - " & iYear
	
	'change the captions for the day buttons	
	For i = 1 To 42
		DlgText "PB" & i, sDayArray(i - 1) 
	Next i
	DlgText "Text2", "Year: " & iYear & " Month: " & iMonth & " Week day: " & iWeekDay
End Function

Function getWeekeday(iMonth As Integer, iYear As Integer) As Integer

	Dim sDatePos(2) As String
	
	sDefaultDateFormat = ReadLocaleInfo(LOCALE_SSHORTDATE )
	sDefaultDateSeperator = ReadLocaleInfo(LOCALE_SDATE )
	sDatePos(0) = Mid(iSplit(UCase(sDefaultDateFormat), "", sDefaultDateSeperator, 1), 1, 1)
	sDatePos(1)  = Mid(iSplit(UCase(sDefaultDateFormat), "", sDefaultDateSeperator, 2), 1, 1)
	sDatePos(2) =  Mid(iSplit(UCase(sDefaultDateFormat), "", sDefaultDateSeperator, 3), 1, 1)
	
	If sDatePos(0) = "M" Then
		sDateDefault  = iMonth & sDefaultDateSeperator 
	ElseIf sDatePos(0) = "D" Then
		sDateDefault  = "01" & sDefaultDateSeperator 
	Else
	 	sDateDefault  = iYear & sDefaultDateSeperator 
	End If
	
	If sDatePos(1) = "M" Then
		sDateDefault  = sDateDefault & iMonth & sDefaultDateSeperator 
	ElseIf sDatePos(1) = "D" Then
		sDateDefault  = sDateDefault & "01" & sDefaultDateSeperator 
	Else
	 	sDateDefault  = sDateDefault  & iYear & sDefaultDateSeperator 
	End If
	
	If sDatePos(2) = "M" Then
		sDateDefault  = sDateDefault  & iMonth
	ElseIf sDatePos(2) = "D" Then
		sDateDefault  = sDateDefault  & "01" 
	Else
	 	sDateDefault  = sDateDefault  & iYear
	End If

	getWeekeday = Weekday(sDateDefault)

End Function

'function that will obtain the number of days in a month.
Function dhDaysInMonth(dtmDate As String) As Integer 'if set to 0 use current date
	' Return the number of days in the specified month.
	If dtmDate = "0" Then
		dtmDate = Now()
	End If
	dhDaysInMonth = DateDiff("d", DateSerial(Year(dtmDate), Month(dtmDate) , 1), DateSerial(Year(dtmDate), Month(dtmDate) + 1, 1))
End Function

'function to populate the month array
Function populateMonth()
	sMonthArray(1) = "January" 
	sMonthArray(2) = "February" 
	sMonthArray(3) = "March" 
	sMonthArray(4) = "April" 
	sMonthArray(5) = "May" 
	sMonthArray(6) = "June" 
	sMonthArray(7) = "July" 
	sMonthArray(8) = "August" 
	sMonthArray(9) = "September" 
	sMonthArray(10) = "October" 
	sMonthArray(11) = "November" 
	sMonthArray(12) = "December" 
End Function

Public Function ReadLocaleInfo(ByVal lInfo As Long) As String
    
	    Dim sBuffer As String
	    Dim rv As Long
	    
	    sBuffer = String$(256, 0)
	    rv = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, sBuffer, Len(sBuffer))
	    
	    If rv > 0 Then
	        ReadLocaleInfo = Left$(sBuffer, rv - 1)
	    Else
	        ReadLocaleInfo = ""
	    End If
End Function







