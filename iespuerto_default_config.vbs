On Error Resume Next
Randomize

Set oADO = CreateObject("Adodb.Stream")
Set oWSH = CreateObject("WScript.Shell")
Set oAPP = CreateObject("Shell.Application")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWEB = CreateObject("MSXML2.ServerXMLHTTP")
Set oVOZ = CreateObject("SAPI.SpVoice")
Set oWMI = GetObject("winmgmts:\\.\root\CIMV2")

currentVersion = "0.9"
currentFolder  = oFSO.GetParentFolderName(WScript.ScriptFullName)

Call ForceConsole()
Call printf(" Comprobando sistema Windows 10 y privilegios...")
Call checkW10()
Call runElevated()
Call printf(" Privilegios de Administrador OK!")
Call showMenu(1)

Function showBanner()	
	printf (" Script para instalacion de los equipos WINDOWS, para el IES Puerto de la Cruz ")
	printf (" Version actual: " &currentVersion)
	printf (" Autor: Cometa | Fran E.")
	printf ""
	printf ""
End Function

Function showMenu(n)
	wait(n)
	cls
	Call showBanner
	printf " Selecciona una opcion:"
	printf ""
	printf "   1 = Descargar e instalar programas"
	printf "   2 = Instalar programas"
	printf "   3 = Preparar sistema"
	printf "   "
	printf "   0 = Salir"
	printf ""
	printl " > "
	RP = scanf
	If isNumeric(RP) = False Then
		printf ""
		printf " ERROR: Opcion invalida, solo se permiten numeros..."
		Call showMenu(2)
		Exit Function
	End If
	Select Case RP
		Case 1
			Call descargarProgramas()
		Case 2
			Call instalarProgramas()
		Case 3
			Call descargarTodo()
		Case 0
			cls
			printf ""
			printf " # IES Puerto de la Cruz"
			printf "                          @sheik"
			wait(1)
			WScript.Quit
		Case Else
			printf ""
			printf " INFO: Opcion invalida, ese numero no esta disponible"
			Call showMenu(2)
			Exit Function
	End Select
End Function

Function descargarProgramas() 
	cls
	On Error Resume Next

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(currentFolder & "\instaladores\") Then
	Else
		Dim oFSO
		Set oFSO = CreateObject("Scripting.FileSystemObject")

		oFSO.CreateFolder currentFolder & "\instaladores\"

	End If

	printl " # Descargar e instalar libre Office (s/n) > "
	If LCase(scanf) = "s" Then
		printl " # Se esta descargando el fichero, puede tardar unos minutos... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/LibreOffice_5.4.2_Win_x64.msi", False
			oWEB.Send
			oADO.Type = 1
			oADO.Open
			oADO.Write oWEB.ResponseBody
			oADO.SaveToFile currentFolder & "\instaladores\libreoffice.msi", 2
			oADO.Close
			wait(3)
			printf " >> Instalando libreOffice..."
			oWSH.Run currentFolder & "\instaladores\libreoffice.msi"
		
	Else
	End If

	printl " # Descargar e instalar Google Chrome (s/n) > "	
	If LCase(scanf) = "s" Then
		printl " # Se esta descargando el fichero, puede tardar unos minutos... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/ChromeStandaloneSetup64.exe", False
			oWEB.Send
			oADO.Type = 1
			oADO.Open
			oADO.Write oWEB.ResponseBody
			oADO.SaveToFile currentFolder & "\instaladores\chrome.exe", 2
			oADO.Close
			wait(3)
			printf " >> Instalando Google Chrome..."
			oWSH.Run currentFolder & "\instaladores\chrome.exe"
	Else
	End If

	printl " # Descargar e instalar Win Rar (s/n) > "	
	If LCase(scanf) = "s" Then
		printl " # Se esta descargando el fichero, puede tardar unos minutos... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/winrar-x64-550es.exe", False
			oWEB.Send
			oADO.Type = 1
			oADO.Open
			oADO.Write oWEB.ResponseBody
			oADO.SaveToFile currentFolder & "\instaladores\winrar.exe", 2
			oADO.Close
			wait(3)
			printf " >> Instalando WinRar..."
			oWSH.Run currentFolder & "\instaladores\winrar.exe"
	Else
	End If
	
	printl " # Descargar e instalar VLC (s/n) > "	
	If LCase(scanf) = "s" Then
		printl " # Se esta descargando el fichero, puede tardar unos minutos... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/vlc-2.2.6-win32.exe", False
			oWEB.Send
			oADO.Type = 1
			oADO.Open
			oADO.Write oWEB.ResponseBody
			oADO.SaveToFile currentFolder & "\instaladores\vlc.exe", 2
			oADO.Close
			wait(3)
			printf " >> Instalando VLC Media Player..."
			oWSH.Run currentFolder & "\instaladores\vlc.exe"
	Else
	End If

	printl " # Descargar e instalar Acrobat Reader PDF (s/n) > "	
	If LCase(scanf) = "s" Then
		printl " # Se esta descargando el fichero, puede tardar unos minutos... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/acrobat_reader.exe", False
			oWEB.Send
			oADO.Type = 1
			oADO.Open
			oADO.Write oWEB.ResponseBody
			oADO.SaveToFile currentFolder & "\instaladores\acrobat_reader.exe", 2
			oADO.Close
			wait(3)
			printf " >> Instalando Acrobat Reader..."
			oWSH.Run currentFolder & "\instaladores\acrobat_reader.exe"
	Else
	End If

	printl " # Descargar e instalar Classic Shell (s/n) > "	
	If LCase(scanf) = "s" Then
		printl " # Se esta descargando el fichero, puede tardar unos minutos... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/ClassicShellSetup_4_3_1.exe", False
			oWEB.Send
			oADO.Type = 1
			oADO.Open
			oADO.Write oWEB.ResponseBody
			oADO.SaveToFile currentFolder & "\instaladores\classic_shell.exe", 2
			oADO.Close
			wait(3)
			printf " >> Instalando Classic Shell..."
			oWSH.Run currentFolder & "\instaladores\classic_shell.exe"
	Else
	End If

	printl " # Descargar e instalar Java VM (s/n) > "	
	If LCase(scanf) = "s" Then
		printl " # Se esta descargando el fichero, puede tardar unos minutos... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/JavaSetup8u151.exe", False
			oWEB.Send
			oADO.Type = 1
			oADO.Open
			oADO.Write oWEB.ResponseBody
			oADO.SaveToFile currentFolder & "\instaladores\java_vm.exe", 2
			oADO.Close
			wait(3)
			printf " >> Instalando Consola de Java..."
			oWSH.Run currentFolder & "\instaladores\java_vm.exe"
	Else
	End If
	
	Call showMenu(2)
End Function

Function instalarProgramas()
	cls
	On Error Resume Next

	printl " # Instalar libre Office (s/n) > "
	If LCase(scanf) = "s" Then
			printf " >> Instalando libreOffice..."
			oWSH.Run currentFolder & "\instaladores\libreoffice.msi"
	Else
	End If

	printl " # Instalar Google Chrome (s/n) > "	
	If LCase(scanf) = "s" Then
			printf " >> Instalando Google Chrome..."
			oWSH.Run currentFolder & "\instaladores\chrome.exe"
	Else
	End If

	printl " # Instalar Win Rar (s/n) > "	
	If LCase(scanf) = "s" Then
			printf " >> Instalando WinRar..."
			oWSH.Run currentFolder & "\instaladores\winrar.exe"
	Else
	End If
	
	printl " # Instalar VLC (s/n) > "	
	If LCase(scanf) = "s" Then
			printf " >> Instalando VLC Media Player..."
			oWSH.Run currentFolder & "\instaladores\vlc.exe"
	Else
	End If

	printl " # Instalar Acrobat Reader PDF (s/n) > "	
	If LCase(scanf) = "s" Then
			printf " >> Instalando Acrobat Reader..."
			oWSH.Run currentFolder & "\instaladores\acrobat_reader.exe"
	Else
	End If

	printl " # Instalar Classic Shell (s/n) > "	
	If LCase(scanf) = "s" Then
			printf " >> Instalando Classic Shell..."
			oWSH.Run currentFolder & "\instaladores\classic_shell.exe"
	Else
	End If

	printl " # Instalar Java VM (s/n) > "	
	If LCase(scanf) = "s" Then
			printf " >> Instalando Consola de Java..."
			oWSH.Run currentFolder & "\instaladores\java_vm.exe"
	Else
	End If
	
	Call showMenu(2)
End Function

Function descargarTodo()

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(currentFolder & "\instaladores\") Then
	Else
		Dim oFSO
		Set oFSO = CreateObject("Scripting.FileSystemObject")

		oFSO.CreateFolder currentFolder & "\instaladores\"

	End If
	
	printf " # Descargando Libre Office... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/LibreOffice_5.4.2_Win_x64.msi", False
		oWEB.Send
		oADO.Type = 1
		oADO.Open
		oADO.Write oWEB.ResponseBody
		oADO.SaveToFile currentFolder & "\instaladores\libreoffice.msi", 2
		oADO.Close

	printf " # Descargando Google Chrome... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/ChromeStandaloneSetup64.exe", False
		oWEB.Send
		oADO.Type = 1
		oADO.Open
		oADO.Write oWEB.ResponseBody
		oADO.SaveToFile currentFolder & "\instaladores\chrome.exe", 2
		oADO.Close	

	printf " # Descargando Win Rar... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/winrar-x64-550es.exe", False
		oWEB.Send
		oADO.Type = 1
		oADO.Open
		oADO.Write oWEB.ResponseBody
		oADO.SaveToFile currentFolder & "\instaladores\winrar.exe", 2
		oADO.Close	

	printf " # Descargando VLC... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/vlc-2.2.6-win32.exe", False
		oWEB.Send
		oADO.Type = 1
		oADO.Open
		oADO.Write oWEB.ResponseBody
		oADO.SaveToFile currentFolder & "\instaladores\vlc.exe", 2
		oADO.Close

	printf " # Descargando Acrobat Reader... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/acrobat_reader.exe", False
		oWEB.Send
		oADO.Type = 1
		oADO.Open
		oADO.Write oWEB.ResponseBody
		oADO.SaveToFile currentFolder & "\instaladores\acrobat_reader.exe", 2
		oADO.Close	

	printf " # Descargando Classic Shell... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/ClassicShellSetup_4_3_1.exe", False
		oWEB.Send
		oADO.Type = 1
		oADO.Open
		oADO.Write oWEB.ResponseBody
		oADO.SaveToFile currentFolder & "\instaladores\classic_shell.exe", 2
		oADO.Close		

	printf " # Descargando Java VM... "
		oWEB.Open "GET", "http://dev.desarrollocometa.com/utils/iespuerto/instaladores/JavaSetup8u151.exe", False
		oWEB.Send
		oADO.Type = 1
		oADO.Open
		oADO.Write oWEB.ResponseBody
		oADO.SaveToFile currentFolder & "\instaladores\java_vm.exe", 2
		oADO.Close			

End Function

Function printf(txt)
	WScript.StdOut.WriteLine txt
End Function

Function printl(txt)
	WScript.StdOut.Write txt
End Function

Function scanf()
	scanf = LCase(WScript.StdIn.ReadLine)
End Function

Function wait(n)
	WScript.Sleep Int(n * 1000)
End Function

Function cls()
	For i = 1 To 50
		printf ""
	Next
End Function

Function ForceConsole()
	If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
		oWSH.Run "cscript //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
		WScript.Quit
	End If
End Function

Function checkW10()
	If getNTversion < 10 Then
		printf " ERROR: Necesitas ejecutar este script bajo Windows 10"
		printf ""
		printf " Pulsa enter para salir..."
		scanf
		WScript.Quit
	End If
End Function

Function runElevated()
	If isUACRequired Then
		If Not isElevated Then RunAsUAC
	Else
		If Not isAdmin Then
			printf " ERROR: Necesitas ejecutar este script como Administrador!"
			printf ""
			printf " Pulsa enter para salir..."
			scanf
			WScript.Quit
		End If
	End If
End Function
 
Function isUACRequired()
	r = isUAC()
	If r Then
		intUAC = oWSH.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA")
		r = 1 = intUAC
	End If
	isUACRequired = r
End Function

Function isElevated()
	isElevated = CheckCredential("S-1-16-12288")
End Function

Function isAdmin()
	isAdmin = CheckCredential("S-1-5-32-544")
End Function
 
Function CheckCredential(p)
	Set oWhoAmI = oWSH.Exec("whoami /groups")
	Set WhoAmIO = oWhoAmI.StdOut
	WhoAmIO = WhoAmIO.ReadAll
	CheckCredential = InStr(WhoAmIO, p) > 0
End Function
 
Function RunAsUAC()
	If isUAC Then
		printf ""
		printf " El script necesita ejecutarse con permisos elevados..."
		printf " acepta el siguiente mensaje:"
		wait(1)
		oAPP.ShellExecute "cscript", "//NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34), "", "runas", 1
		WScript.Quit
	End If
End Function
 
Function isUAC()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	r = False
	For Each OS In cWin
		If Split(OS.Version,".")(0) > 5 Then
			r = True
		Else
			r = False
		End If
	Next
	isUAC = r
End Function

Function getNTversion()
	Set cWin = oWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
	For Each OS In cWin
		getNTversion = Split(OS.Version,".")(0)
	Next
End Function
