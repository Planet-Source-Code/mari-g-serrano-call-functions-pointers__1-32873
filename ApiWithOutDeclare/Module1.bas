Attribute VB_Name = "mCallFunct"
Option Explicit
'  MaRiØ Glez Serrano. 02/02
' Llama a funciones sin declararlas ( _stdcall y cDecl ) (no se si tambien Alpha)
' CallApiByName llama a una funcion pasandole la dll y el nombre de la funcion
' CallPtr llama a una funcion (de otro proceso) pasandole la direccion
' '¡cuidadin con pasar datos q no espera! o.. (ya sabes...CrasH!!)

' Call Functions Without Declares!(_stdcall and cDecl)
' CallApiByName call a function pasing the dll name and the function name
' CallPtr calls a function passing the address of the fuuction
' be careful with wrong parameter types..(you know...crash!)

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
                         (ByVal lpLibFileName As String) As Long
Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                          ByVal lpProcName As String) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                           (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
                            ByVal Msg As Long, ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                    (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
                    
Private mlngParameters() As Long ' lista de parámetros//parameters list
Private mlngAddress As Long      ' direccion de la funcion a llamar//functioon pointer to call
Private mbytCode() As Byte       ' buffer para codigo en ensamblador//buffer to ASM code
Private mlngCP As Long           ' para saber donde está  //to store where are
                                 '  el ultimo byte  // the last byte
                                 ' añadido al codigo//added to the code
                       
   Dim ASMCode() As Byte 'To Test a call to ASM Code
  
Public Function CallApiByName(libName As String, funcName As String, _
                              ParamArray lngFuncParams()) As Long
   'IN: Nombre dll, Nombre Funcion (usar nombre original (el Alias)),parametros (Longs!!)
   'OUT:long
   'si se pasan longs,ej: x = CallApiByName("user32", "FlashWindow", hWnd, 1&)
   '·si se quiere pasar un string(ANSI) pasar el puntero a un array de bytes vbFromUnicode
   'ej:Dim s() As Byte
   '   s = StrConv("MaRiO", vbFromUnicode)
   '   (FuncParams:=VarPtr(s(0))
   '·Si se quiere pasar un String Unicode, pasar el puntero al string
   'para el resto de cosas q no sean Longs, pasar el puntero.
   
   'IN: dll Name, Function Name (use original name(Alias)),pass longs parameters!!
   'OUT:long
   '·you can pass longs,eg: x = CallApiByName("user32", "FlashWindow", hWnd, 1&)
   '·to pass a string(ANSI) pass the pointer to an byte array (vbFromUnicode)
   ' eg:Dim s() As Byte
   '    s = StrConv("MaRiO", vbFromUnicode)
   '   (FuncParams:=VarPtr(s(0))
   '·to pass an Unicode string , pass the pointer to string
   'pass UDT´s... pass te object pointer
   
On Error GoTo errTipos
   Dim lb As Long, i As Integer
   ReDim mlngParameters(0)
   ReDim mbytCode(0)
   mlngAddress = 0
   lb = LoadLibrary(ByVal libName)
   If lb = 0 Then
      MsgBox "DLL not Found.", vbCritical 'Dll no encontrada o No cargada correctamente
      Exit Function
   End If
   mlngAddress = GetProcAddress(lb, ByVal funcName)
   If mlngAddress = 0 Then
      MsgBox "Function not found in " & libName, vbCritical 'Funcion no encontrada en
      FreeLibrary lb
      Exit Function
   End If
   ReDim mlngParameters(UBound(lngFuncParams) + 1)
   For i = 1 To UBound(mlngParameters)
      mlngParameters(i) = CLng(lngFuncParams(i - 1))
   Next i
   CallApiByName = CallWindowProc(PrepareCode, 0, 0, 0, 0)
   FreeLibrary lb
   Exit Function
errTipos:

    MsgBox Err.Description & "Pass only Longs!" 'Pasa solo Longs!
    FreeLibrary lb
End Function

Public Function CallPtr(lPtr As Long, ParamArray lngFuncParams()) As Long
    ' usa la funcion 'Address' junto con el Operador AddressOf para
    ' asignar la direccion a una variable

    ' Use the 'Address' function with the AddressOf operator
    ' to assign the address to a variable
    
    Dim i As Long
    mlngAddress = lPtr
    ReDim mlngParameters(UBound(lngFuncParams) + 1)
    For i = 1 To UBound(mlngParameters)
       mlngParameters(i) = CLng(lngFuncParams(i - 1))
    Next i
    CallPtr = CallWindowProc(PrepareCode, 0, 0, 0, 0)
    
End Function

Private Function PrepareCode() As Long
    Dim lngX As Long, codeStart As Long
    ReDim mbytCode(18 + 32 + 6 * UBound(mlngParameters))
    codeStart = GetAlignedCodeStart(VarPtr(mbytCode(0)))
    mlngCP = codeStart - VarPtr(mbytCode(0))
    For lngX = 0 To mlngCP - 1
        mbytCode(lngX) = &HCC
    Next
    'ASM Code
    AddByteToCode &H58 'pop eax
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H59 'pop ecx
    AddByteToCode &H50 'push eax
    For lngX = UBound(mlngParameters) To 1 Step -1
        AddByteToCode &H68 'push wwxxyyzz
        AddLongToCode mlngParameters(lngX)
    Next
    AddCallToCode mlngAddress 'function Pointer
    AddByteToCode &HC3
    AddByteToCode &HCC
    PrepareCode = codeStart
End Function

Private Sub AddCallToCode(lngAddress As Long)
    AddByteToCode &HE8
    AddLongToCode lngAddress - VarPtr(mbytCode(mlngCP)) - 4
End Sub

Private Sub AddLongToCode(lng As Long)
    Dim intX As Integer
    Dim byt(3) As Byte
    CopyMemory byt(0), lng, 4
    For intX = 0 To 3
        AddByteToCode byt(intX)
    Next
End Sub

Private Sub AddByteToCode(byt As Byte)
    mbytCode(mlngCP) = byt
    mlngCP = mlngCP + 1
End Sub

Private Function GetAlignedCodeStart(lngAddress As Long) As Long
    GetAlignedCodeStart = lngAddress + (15 - (lngAddress - 1) Mod 16)
    If (15 - (lngAddress - 1) Mod 16) = 0 Then
        GetAlignedCodeStart = GetAlignedCodeStart + 16
    End If
End Function

Public Function Address(lPtr As Long) As Long
    'IN: AddressOf a Function.OUT: la address in Long
    Address = CLng(lPtr)
End Function


Public Function Test(ByVal x As Long) As Long
    Test = x + 1
End Function

Public Function TestMsg() As Long
    TestMsg = MsgBox("This Function was Called by his Pointer")
End Function



Public Function CallASM() As Long
   '     mov   ecx, [esp + 4]
   '     8b 4c 24 04
   '     mov   eax, [ebp + ecx]
   '     8b 44 0d 00
   '     ret   4
   '     c2 04 00
   'maybe this don´t work..
   ReDim ASMCode(10) As Byte
   ASMCode(0) = &H8B
   ASMCode(1) = &H4C
   ASMCode(2) = &H24
   ASMCode(3) = &H4
   
   ASMCode(4) = &H8B
   ASMCode(5) = &H44
   ASMCode(6) = &HD
   ASMCode(7) = &H0
   
   ASMCode(8) = &HC2
   ASMCode(9) = &H4
   ASMCode(10) = &H0
   
   Dim x As Long
   x = GetAlignedCodeStart(VarPtr(ASMCode(0)))
   CallASM = CallWindowProc(x, 0, 0, 0, 0)
End Function


Public Function R() As Long
    R = 13
End Function
