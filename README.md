<div align="center">

## RegCodes


</div>

### Description

This class contains two functions which can be helpful in creating an online shareware registration system for your software projects. GenerateKeyCode takes a username, or any other string, and generates a unique human-readable registration code (such as 9397-JQM0LD0YJV from the string: Andy Carrasco). GenerateKeyCode will generate a totally unique registration code over and over again, even for the exact same name! VerifyKeyCode is the partner function, and will verify if a keycode matches a given name.
 
### More Info
 
IMPORTANT NOTE!

Although the codes generated from this algorithm will throughly confuse, and secure your code from, the average user, I make absolutely no gaurantee of security. The average hacker is NOT the average user, and anyone with a fairly general understanding of cyphering could quickly crack these algorithms. On the other hand, there are NO registration code utilities which gaurantee security, it would be foolish to believe that any form of encryption is totally secure. You may freely, and are encouraged to, use this algorithm in your own registration utilities, provided that you fully understand that I do not gaurantee the security of these functions, and that I will take no liability for any losses occuring from your use of these functions. They are primarily intended as a learning facility.

Andy Carrasco


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andy Carrasco](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andy-carrasco.md)
**Level**          |Unknown
**User Rating**    |6.0 (615 globes from 103 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andy-carrasco-regcodes__1-1199/archive/master.zip)





### Source Code

```
Option Explicit
' Name: GenerateKeyCode
'
' Description:
'  This little routine generates a keycode for shareware registration in the
'  format XXXX-YYYYYYYYYY, based on the Name given as an argument. The first
'  four digits are a randomly generated seed value, which makes 8999 possible keycodes
'  for people with the same name (like John Smith). The last four digits are
'  the actual code.
'
' Written by:
'  Andy Carrasco (Copyright 1998)
'
Public Function GenerateKeyCode(sName As String) As String
  Dim sRandomSeed As String
  Dim sKeyCode As String
  Dim X As Long
  Dim KeyCounter As Long
  Dim PrimaryLetter As Long
  Dim CodedLetter As Long
  Dim sBuffer As String
  Randomize
  sRandomSeed = CStr(Int((9999 - 1000 + 1) * Rnd + 1000))
  sName = UCase$(sName)
  KeyCounter = 1
  'Clean up sName so there are no illegal characters.
  For X = 1 To Len(sName)
    If Asc(Mid$(sName, X, 1)) >= 65 And Asc(Mid$(sName, X, 1)) <= 90 Then sBuffer = sBuffer & Mid$(sName, X, 1)
  Next X
  sName = sBuffer
  'if the name is less than 10 characters long, pad it out with ASCII 65
  Do While Len(sName) < 10
    sName = sName + Chr$(65)
  Loop
  For X = 1 To Len(sName)
    PrimaryLetter = Asc(Mid$(sName, X, 1))
    CodedLetter = PrimaryLetter + CInt(Mid$(sRandomSeed, KeyCounter, 1))
    If CodedLetter < 90 Then
      sKeyCode = sKeyCode + Chr$(CodedLetter)
    Else
      sKeyCode = sKeyCode + "0"
    End If
    'Increment the keycounter
    KeyCounter = KeyCounter + 1
    If KeyCounter > 4 Then KeyCounter = 1
  Next X
  GenerateKeyCode = sRandomSeed + "-" + Left$(sKeyCode, 10)
End Function
' Name: VerifyKeyCode
'
' Description:
'  Verifies if a given keycode is valid for a given name.
'
' Parameters:
'  sName  - A string containing the user name to validate the key against
'  sKeyCode- A string containins the keycode in the form XXXX-YYYYYYYYYY.
'
Public Function VerifyKeyCode(sName As String, sKeyCode As String) As Boolean
  Dim sRandomSeed As String
  Dim X As Long
  Dim KeyCounter As Long
  Dim PrimaryLetter As Long
  Dim DecodedKey As String
  Dim AntiCodedLetter As Long
  Dim sBuffer As String
  sRandomSeed = Left$(sKeyCode, InStr(sKeyCode, "-") - 1)
  sName = UCase$(sName)
  sKeyCode = Right$(sKeyCode, 10)
  KeyCounter = 1
  'Clean up sName so there are no illegal characters.
  For X = 1 To Len(sName)
    If Asc(Mid$(sName, X, 1)) >= 65 And Asc(Mid$(sName, X, 1)) <= 90 Then sBuffer = sBuffer & Mid$(sName, X, 1)
  Next X
  sName = sBuffer
  'if the name is less than 10 characters long, pad it out with ASCII 65
  Do While Len(sName) < 10
    sName = sName + Chr$(65)
  Loop
  'now, decode the keycode
  For X = 1 To Len(sKeyCode)
    PrimaryLetter = Asc(Mid$(sKeyCode, X, 1))
    AntiCodedLetter = PrimaryLetter - CInt(Mid$(sRandomSeed, KeyCounter, 1))
    If PrimaryLetter = 48 Then 'zero
      DecodedKey = DecodedKey + Mid$(sName, X, 1) 'Take the corresponding letter from the name
    Else
      DecodedKey = DecodedKey + Chr$(AntiCodedLetter)
    End If
    'Increment the keycounter
    KeyCounter = KeyCounter + 1
    If KeyCounter > 4 Then KeyCounter = 1
  Next X
  If DecodedKey = Left$(sName, 10) Then
    VerifyKeyCode = True
  Else
    VerifyKeyCode = False
  End If
End Function
```

