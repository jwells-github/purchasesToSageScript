#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.


# Pressing CTRL and J at the same time will start the script
# BEFORE STARTING
#   The suppliers name should be selected in Excel 
#   Both the invoice and credit form should be open in Sage
#   Excel should be filtered to one currency (GBP or USD)

^j::

# The script will repeat until there are no more invoices / credits left
Loop {	

	#Copies the invoice / credit value to the clipboard
	Send {Right}{Right}

	Sleep, 150

	
	Send {F2}
	Send +{Home}
	Send ^c

	Sleep, 150	

	#Copies the supplier name to the clipboard
 	Send {Esc}
   	Send {Left}{Left}

   	Sleep, 150

   	Send {F2}
   	Send +{Home}
   	Send ^c

   	Sleep, 150
   	WinActivate, Batch Supplier Invoice	
   	Sleep, 150
   	#Ends the script if the clipboard is empty
   	if (Clipboard = "")
	{
		break
   	}
	#Sends the relevant supplier code to Sage
   	else if (Clipboard = "Supplier 1")
	{
		Send SUP1
   	}
	else if (Clipboard = "Supplier 2")
	{
		Send SUP2
   	}
	else if (Clipboard = "Supplier 3")
	{
		Send SUP3
   	}
	else if (Clipboard = "Supplier 4")
	{
		Send SUP4
   	}
	else if (Clipboard = "Supplier 5")
	{
		Send SUP5
   	}
	#If an unrecognised supplier name is copied, the script ends and alerts the user
	else 
	{
		MsgBox, Unrecognised supplier - script stopped
		ExitApp
	}
   	Sleep, 150
   	WinActivate Purchases - Excel
   	Sleep, 150

   	#Copies the invoice date to the clipboard
   	Send {Esc}
   	Send {Right}{Right}{Right}{Right}{Right}

   	Sleep, 150

   	Send {F2}
   	Send +{Home}
   	Send ^c


   	Sleep, 150
   	WinActivate, Batch Supplier Invoice	
   	Sleep, 150

   	#Pastes the invoice date into Sage
   	Send {Tab}
   	Sleep, 150
   	SendRaw, %Clipboard%

   	Sleep, 150
   	WinActivate Purchases - Excel
   	Sleep, 150

   	#Copies the invoice number to the clipboard
   	Send {Esc}
   	Send {Left}{Left}{Left}{Left}{Left}{Left}{Left}

   	Sleep, 150

   	Send {F2}
   	Send +{Home}
   	Send ^c

   	Sleep, 150
   	WinActivate, Batch Supplier Invoice	
   	Sleep, 150


   	#Pastes the invoice number into Sage
   	Send {Tab}
   	Sleep, 150
   	Send P
   	SendRaw, %Clipboard%

   	Sleep, 150
   	WinActivate Purchases - Excel
   	Sleep, 150

   	#Copies the supplier's invoice number to the clipboard
   	Send {Esc}
   	Send {Right}{Right}{Right}{Right}{Right}{Right}{Right}{Right}

   	Sleep, 150

   	Send {F2}
   	Send +{Home}
   	Send ^c

   	Sleep, 150
   	WinActivate, Batch Supplier Invoice	
   	Sleep, 150

   	#Pastes the supplier's invoice number into Sage
  	Send {Tab}
   	Sleep, 150
   	SendRaw, %Clipboard%

   	Sleep, 150
   	WinActivate Purchases - Excel
   	Sleep, 150


   	#Copies our company reference to the clipboard
   	Send {Esc}
   	Send {Right}
    
   	Sleep, 150

   	Send {F2}
   	Send +{Home}
   	Send ^c

   	Sleep, 150
   	WinActivate, Batch Supplier Invoice	
   	Sleep, 150

   	#Pastes our company reference into Sage
   	Send {Tab}{Tab}{Tab}{Tab}{Tab}
   	Sleep, 150
   	SendRaw, %Clipboard%

   	Sleep, 150
   	WinActivate Purchases - Excel
   	Sleep, 150

   	#Copies the invoice value to the clipboard
   	Send {Esc}
   	Send {Left}{Left}{Left}{Left}{Left}

   	Sleep, 150

   	Send {F2}
   	Send +{Home}
   	Send ^c

   	Sleep, 150
   	WinActivate, Batch Supplier Invoice	
   	Sleep, 150

   	#Pastes the invoice value into Sage
   	Send {Tab}
   	Sleep, 150
   	SendRaw, %Clipboard%

   	Sleep, 150
   	Send {Tab}{Tab}{Tab}
   	Clipboard = ;

    Sleep, 150
   	WinActivate Purchases - Excel
   	Sleep, 150

   	#Deselects the current cell and moves onto the next invoice / credit
	Send {Esc}
   	Send {Left}{Left}
	Send {Down}
	Sleep, 150

}
Return