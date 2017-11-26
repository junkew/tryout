'***********************************************************************************************
'   * Please leave any Trademarks or Credits in place.
'   *
'   * ACKNOWLEDGEMENT TO CONTRIBUTORS :
'   *       STEPHEN BULLEN, 15 November 1998 - Original PastPicture code
'   *       G HUDSON, 5 April 2010 - Pause Function
'   *       LUTZ GENTKOW, 23 July 2011 - Alt + PrtScrn
'   *       PAUL FRANCIS, 11 April 2013 - Putting all pieces together, bridging the 32 bit and 64 bit version.
'   *       CHRIS O, 12 April 2013 - Code suggestion to work on older versions of Access.
'   *
'   * DESCRIPTION: Creates a standard Picture object from whatever is on the clipboard.
'   *              This object is then saved to a location on the disc. Please note, this
'   *              can also be assigned to (for example) and Image control on a userform.
'   *
'   * The code requires a reference to the "OLE Automation" type library.
'   *
'   * The code in this module has been derived from a number of sources
'   * discovered on MSDN, Access World Forum, VBForums.
'   *
'   * To use it, just copy this module into your project, then you can use:
'   * SaveClip2Bit("C:\Pics\Sample.bmp")
'   * to save this to a location on the Disc.
'   * (Or)
'   * Set ImageControl.Image = PastePicture
'   * to paste a picture of whatever is on the clipboard into a standard image control.
'   *
'   * PROCEDURES:
'   *   PastePicture  :   The entry point for 'Setting' the Image
'   *   CreatePicture :   Private function to convert a bitmap or metafile handle to an OLE reference
'   *   fnOLEError    :   Get the error text for an OLE error code
'   *   SaveClip2Bit  :   The entry point for 'Saving' the Image, calls for PastePicture
'   *   AltPrintScreen:   Performs the automation of Alt + PrtScrn, for getting the Active Window.
'   *   Pause         :   Makes the program wait, to make sure proper screen capture takes place.
'**************************************************************************************************

Option Explicit
Option Compare Text

'Declare a UDT to store a GUID for the IPicture OLE Interface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Declare a UDT to store the bitmap information
Private Type uPicDesc
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

'Windows API Function Declarations
#If Win64 = 1 And VBA7 = 1 Then
    
    'Does the clipboard contain a bitmap/metafile?
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
    
    'Open the clipboard to read
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    
    'Get a pointer to the bitmap/metafile
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
    
    'Close the clipboard
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    
    'Convert the handle into an OLE IPicture interface.
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    
    'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
    Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
    
    'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
    Declare PtrSafe Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
    'Uses the Keyboard simulation
    Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

#Else

    'Does the clipboard contain a bitmap/metafile?
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
    
    'Open the clipboard to read
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    
    'Get a pointer to the bitmap/metafile
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
    
    'Close the clipboard
    Private Declare Function CloseClipboard Lib "user32" () As Long
    
    'Convert the handle into an OLE IPicture interface.
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (PicDesc As uPicDesc, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
    
    'Create our own copy of the metafile, so it doesn't get wiped out by subsequent clipboard updates.
    Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
    
    'Create our own copy of the bitmap, so it doesn't get wiped out by subsequent clipboard updates.
    Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    
    'Uses the Keyboard simulation
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

#End If
  
'The API format types we're interested in
Const CF_BITMAP = 2
Const CF_PALETTE = 9
Const CF_ENHMETAFILE = 14
Const IMAGE_BITMAP = 0
Const LR_COPYRETURNORG = &H4

Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SNAPSHOT = &H2C
Private Const VK_MENU = &H12

' Subroutine    : AltPrintScreen
' Purpose       : Capture the Active window, and places on the Clipboard.

Sub AltPrintScreen()
    keybd_event VK_MENU, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
End Sub

' Subroutine    : PastePicture
' Purpose       : Get a Picture object showing whatever's on the clipboard.

Function PastePicture() As IPicture
    'Some pointers
    Dim h As Long, hPtr As Long, hPal As Long, lPicType As Long, hCopy As Long

    'Check if the clipboard contains the required format
    If IsClipboardFormatAvailable(CF_BITMAP) Then
        'Get access to the clipboard
        h = OpenClipboard(0&)
        If h > 0 Then
            'Get a handle to the image data
            hPtr = GetClipboardData(CF_BITMAP)

            hCopy = CopyImage(hPtr, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)

            'Release the clipboard to other programs
            h = CloseClipboard
            'If we got a handle to the image, convert it into a Picture object and return it
            If hPtr <> 0 Then Set PastePicture = CreatePicture(hCopy, 0, CF_BITMAP)
        End If
    End If
End Function


' Subroutine    : CreatePicture
' Purpose       : Converts a image (and palette) handle into a Picture object.
' NOTE          : Requires a reference to the "OLE Automation" type library

Private Function CreatePicture(ByVal hPic As Long, ByVal hPal As Long, ByVal lPicType) As IPicture
    ' IPicture requires a reference to "OLE Automation"
    Dim r As Long, uPicInfo As uPicDesc, IID_IDispatch As GUID, IPic As IPicture
    'OLE Picture types
    Const PICTYPE_BITMAP = 1
    Const PICTYPE_ENHMETAFILE = 4
    ' Create the Interface GUID (for the IPicture interface)
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    ' Fill uPicInfo with necessary parts.

    With uPicInfo
        .Size = Len(uPicInfo) ' Length of structure.
        .Type = PICTYPE_BITMAP ' Type of Picture
        .hPic = hPic ' Handle to image.
        .hPal = hPal ' Handle to palette (if bitmap).
    End With

    ' Create the Picture object.
    r = OleCreatePictureIndirect(uPicInfo, IID_IDispatch, True, IPic)

    ' If an error occured, show the description
    If r <> 0 Then Debug.Print "Create Picture: " & fnOLEError(r)

    ' Return the new Picture object.
    Set CreatePicture = IPic
End Function


' Subroutine    : fnOLEError
' Purpose       : Gets the message text for standard OLE errors

Private Function fnOLEError(lErrNum As Long) As String
    'OLECreatePictureIndirect return values
    Const E_ABORT = &H80004004
    Const E_ACCESSDENIED = &H80070005
    Const E_FAIL = &H80004005
    Const E_HANDLE = &H80070006
    Const E_INVALIDARG = &H80070057
    Const E_NOINTERFACE = &H80004002
    Const E_NOTIMPL = &H80004001
    Const E_OUTOFMEMORY = &H8007000E
    Const E_POINTER = &H80004003
    Const E_UNEXPECTED = &H8000FFFF
    Const S_OK = &H0

    Select Case lErrNum
        Case E_ABORT
            fnOLEError = " Aborted"
        Case E_ACCESSDENIED
            fnOLEError = " Access Denied"
        Case E_FAIL
            fnOLEError = " General Failure"
        Case E_HANDLE
            fnOLEError = " Bad/Missing Handle"
        Case E_INVALIDARG
            fnOLEError = " Invalid Argument"
        Case E_NOINTERFACE
            fnOLEError = " No Interface"
        Case E_NOTIMPL
            fnOLEError = " Not Implemented"
        Case E_OUTOFMEMORY
            fnOLEError = " Out of Memory"
        Case E_POINTER
            fnOLEError = " Invalid Pointer"
        Case E_UNEXPECTED
            fnOLEError = " Unknown Error"
        Case S_OK
            fnOLEError = " Success!"
    End Select
End Function

' Routine   : SaveClip2Bit
' Purpose   : Saves Picture object to desired location.
' Arguments : Path to save the file

Public Sub SaveClip2Bit(savePath As String)
On Error GoTo errHandler:
    AltPrintScreen
    Pause (3)
    SavePicture PastePicture, savePath
errExit:
        Exit Sub
errHandler:
    Debug.Print "Save Picture: (" & err.Number & ") - " & err.Description
    Resume errExit
End Sub

' Routine   : Pause
' Purpose   : Gives a short intreval for proper image capture.
' Arguments : Seconds to wait.

Public Function Pause(NumberOfSeconds As Variant)
On Error GoTo Err_Pause
    Dim PauseTime As Variant, start As Variant
    PauseTime = NumberOfSeconds
    start = Timer
    Do While Timer < start + PauseTime
        DoEvents
    Loop
Exit_Pause:
    Exit Function
Err_Pause:
    MsgBox err.Number & " - " & err.Description, vbCritical, "Pause()"
    Resume Exit_Pause
End Function
