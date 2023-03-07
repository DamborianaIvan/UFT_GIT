Option Explicit

Call IniExe()
	Dim program_route, txt_usuario, txt_securepassword, start_time, stop_time, numIter
	
	numIter=Environment.Value("ActionIteration")
	program_route = "C:\Program Files (x86)\Micro Focus\UFT One\samples\Flights Application\FlightsGUI.exe"
	txt_usuario = DataTable("i_User", dtLocalSheet)
	txt_securepassword = DataTable("i_PasswordSecure", dtLocalSheet)

	
Sub Login_Flow()
	
	start_time = Timer
	SystemUtil.Run program_route
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Open_App.png", True
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set txt_usuario @@ hightlight id_;_1952350608_;_script infofile_;_ZIP::ssf2.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Entry_User.png", True
	WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure txt_securepassword @@ hightlight id_;_1952350944_;_script infofile_;_ZIP::ssf4.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Entry_Password.png", True
	WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click
	While WpfWindow("Micro Focus MyFlight Sample").WpfObject("Hello").Exist = False
		wait 1
		If WpfWindow("Micro Focus MyFlight Sample").Dialog("Login Failed").Exist Then
		   DataTable("o_ValLoginName", dtLocalSheet) = "Login Filed: Change User or Password"
		   WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Login_Failed.png", True
		   imagenToWord "X.- Login Filed", RutaEvidencias() & numIter&".Login_Failed.png"
		   WpfWindow("Micro Focus MyFlight Sample").Dialog("Login Failed").WinButton("Aceptar").Click
		   WpfWindow("Micro Focus MyFlight Sample").Close @@ hightlight id_;_331398_;_script infofile_;_ZIP::ssf28.xml_;_
		   stop_time = Timer
		   DataTable("Timer", dtLocalsheet) = "Se Ejecutó en: ["&stop_time - start_time&"] segundos."
		   Reporter.ReportEvent micFail, "Login Failed", "Change User or Password"
		   ExitTest
		End If
	Wend
	wait 1
	WpfWindow("Micro Focus MyFlight Sample").WpfObject("Maria Johnson").Output CheckPoint("NameUser") @@ hightlight id_;_1921298472_;_script infofile_;_ZIP::ssf28.xml_;_
	WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap RutaEvidencias() & numIter&".Validate_Login.png", True
	stop_time = Timer
	DataTable("Timer", dtLocalsheet) = "Se Ejecutó en: ["&stop_time - start_time&"] segundos."
End Sub
Sub ImagenWord()
	imagenToWord "1.- Open App MyFlight Sample", RutaEvidencias() & numIter&".Open_App.png"
	imagenToWord "2.- Entry User", RutaEvidencias() & numIter&".Entry_User.png"
	imagenToWord "3.- Entry Password", RutaEvidencias() & numIter&".Entry_Password.png"
	imagenToWord "4.- Validate Login", RutaEvidencias() & numIter&".Validate_Login.png"
End Sub

Call Login_Flow()
Call ImagenWord()


