Option Explicit

Dim From_City, To_City, xDate, xClass, Num_Pass, Pass_Name, start_time, stop_time, numIter

		numIter=Environment.Value("ActionIteration")
		From_City = DataTable("i_FromCity", dtLocalSheet)
		To_City = DataTable("i_ToCity", dtLocalSheet)
		xDate = DataTable("i_Date", dtLocalSheet)
		xClass = DataTable("i_Class", dtLocalSheet)
		Num_Pass = DataTable("i_NumPass", dtLocalSheet)
		Pass_Name = DataTable("i_PassName", dtLocalSheet)

Sub Entry_Data()
	start_time = Timer
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select From_City
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select To_City @@ hightlight id_;_1954499544_;_script infofile_;_ZIP::ssf10.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage_2").Click 10,14 @@ hightlight id_;_2076480008_;_script infofile_;_ZIP::ssf11.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("DO").SetDate xDate @@ hightlight id_;_2076482504_;_script infofile_;_ZIP::ssf12.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select xClass @@ hightlight id_;_1952354016_;_script infofile_;_ZIP::ssf16.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select Num_Pass @@ hightlight id_;_2076483704_;_script infofile_;_ZIP::ssf20.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Entry_Data.png", True
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click
	If WpfWindow("Micro Focus MyFlight Sample").Dialog("Error").Exist Then
		   DataTable("o_NumOrder", dtLocalSheet) = "Error: Flight date must be later than today"
		   WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".DATE_Failed.png", True
		   imagenToWord "X.- DATE Filed", RutaEvidencias() & numIter&".DATE_Failed.png"
		   WpfWindow("Micro Focus MyFlight Sample").Dialog("Error").WinButton("Aceptar").Click
		   stop_time = Timer
		   DataTable("Timer", dtLocalsheet) = "Se Ejecutó en: ["&stop_time - start_time&"] segundos."
		   Reporter.ReportEvent micWarning, "Date Failed", "Error: Flight date must be later than today"
		   ExitActionIteration
	End If
End Sub
Sub Select_Flight()
	WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,0
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Select_Flight.png", True
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set Pass_Name
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Resume_Order.png", True
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click	
End Sub @@ hightlight id_;_1954435560_;_script infofile_;_ZIP::ssf21.xml_;_
Sub Wait_OrderAndCapture()
	
 	While WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage_5").Exist = False
		wait 1	
	 Wend
	 wait 2 @@ hightlight id_;_1915348952_;_script infofile_;_ZIP::ssf29.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").WpfObject("$138.60").Output CheckPoint("Ticket_Price")
	WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 137 completed").Output CheckPoint("Order Generate") @@ hightlight id_;_66142_;_script infofile_;_ZIP::ssf28.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Generate_Order.png", True
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("NEW SEARCH").Click
	stop_time = Timer
	DataTable("Timer", dtLocalsheet) = "Se Ejecutó en: ["&stop_time - start_time&"] segundos."
End Sub @@ hightlight id_;_1965591312_;_script infofile_;_ZIP::ssf31.xml_;_
Sub ImagenWord()
	imagenToWord "5.- Data Entry", RutaEvidencias() & numIter&".Entry_Data.png"
	imagenToWord "6.- Select Flight", RutaEvidencias() & numIter&".Select_Flight.png"
	imagenToWord "7.- Resume Order", RutaEvidencias() & numIter&".Resume_Order.png"
	imagenToWord "8.- Generate Order", RutaEvidencias() & numIter&".Generate_Order.png"
End Sub

Call Entry_Data()
Call Select_Flight()
Call Wait_OrderAndCapture()
Call ImagenWord()

