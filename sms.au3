#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <File.au3>

; send over com port
Func SendData($port, $data)
    Local $hFile = FileOpen($port, 2)
    If $hFile = -1 Then
        MsgBox($MB_OK, "Error", "Failed to open COM port.")
        Return False
    EndIf
    FileWrite($hFile, $data)
    FileClose($hFile)
    Return True
EndFunc

; read com port data
Func ReadData($port, $timeout = 5000)
    Local $hFile = FileOpen($port, 0)
    If $hFile = -1 Then
        MsgBox($MB_OK, "Error", "Failed to open COM port.")
        Return ""
    EndIf
    Local $startTime = TimerInit()
    Local $sResponse = ""
    While TimerDiff($startTime) < $timeout
        If FileGetMsg($hFile) <> "" Then
            $sResponse &= FileRead($hFile)
        EndIf
        Sleep(100) ; small delay, helps reduce cpu usage
    WEnd
    FileClose($hFile)
    Return $sResponse
EndFunc

; populate com ports dropdown
Func PopulateCOMPorts()
    Local $sPortList[1] = ["Select COM Port"]
    Local $iIndex = 1

    For $i = 1 To 255
        Local $sPort = "COM" & $i
        If FileExists("\\.\\" & $sPort) Then
            ReDim $sPortList[$iIndex + 1]
            $sPortList[$iIndex] = $sPort
            $iIndex += 1
        EndIf
    Next

    GUICtrlSetData($comPortDropdown, $sPortList)
EndFunc

Func SendSMS()
    Local $comPort = GUICtrlRead($comPortDropdown)
    Local $phoneNumber = GUICtrlRead($phoneNumberInput)
    Local $message = GUICtrlRead($messageInput)

    If StringStripWS($phoneNumber, 3) = "" Or StringStripWS($message, 3) = "" Then
        MsgBox($MB_OK, "Error", "Please enter a phone number and a message.")
        Return
    EndIf

    ; send AT+CMGF=1 command to set sms mode
    If SendData($comPort, "AT+CMGF=1" & @CRLF) Then
        Sleep(2000) ; 2 second delay to ensure module has done command
    EndIf

    ; set recipient's phone number
    If SendData($comPort, 'AT+CMGS="' & $phoneNumber & '"' & @CRLF) Then
        Sleep(2000) ; delay before sending message
    EndIf

    ; Send the message followed by Ctrl+Z character
    If SendData($comPort, $message & Chr(26)) Then
        Sleep(5000) ; wait for message to be processed and sent
    EndIf

    ; read response from gsm module
    Local $response = ReadData($comPort)
    If $response <> "" Then
        MsgBox($MB_OK, "Response", $response) ; display raw response for debugging
    EndIf

    MsgBox($MB_OK, "Success", "Message sent successfully!")
EndFunc

; gui setup
GUICreate("SMS Sender", 320, 250)
GUICtrlCreateLabel("COM Port:", 10, 10, 100, 20)
$comPortDropdown = GUICtrlCreateCombo("", 120, 10, 180, 20)
GUICtrlCreateLabel("Phone Number:", 10, 40, 100, 20)
$phoneNumberInput = GUICtrlCreateInput("", 120, 40, 180, 20)
GUICtrlCreateLabel("Message:", 10, 70, 100, 20)
$messageInput = GUICtrlCreateInput("", 120, 70, 180, 100)
$sendButton = GUICtrlCreateButton("Send SMS", 10, 180, 290, 30)
GUISetState()

; populate com ports dropdown
PopulateCOMPorts()

While 1
    $msg = GUIGetMsg()
    If $msg = $GUI_EVENT_CLOSE Then ExitLoop
    If $msg = $sendButton Then
        ; disable button, prevent multiple clicks
        GUICtrlSetState($sendButton, $GUI_DISABLE)
        SendSMS()
        ; re-enable button after sending
        GUICtrlSetState($sendButton, $GUI_ENABLE)
    EndIf
WEnd
