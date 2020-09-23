Attribute VB_Name = "Parsing_Module"
  Option Explicit

' ******************************************************************************
' ******************************************************************************
' (c) NeoProgrammics 2001 - Jay Tanner
'
' PARSING MODULE
'
' MODULE VERSION: 2001.0624.0232
' G5 Series
' ******************************************************************************
' EXTERNAL DEPENDENCIES: NONE
'
'
' ------------------------------------------------------------------------------
' FUNCTIONS INDEX:
'
'
' Num_Part_Of(ArgString)
'
' Fetches only the numerical part from a single compound data argument string.
' It is the counterpart of the Text_Part_Of() function below.
'
' ------------------------------------------------------------------------------
' Num_Val_Of(ArgString)
'
' Evaluates a single numerical string argument like that returned by the above
' function.  The (ArgString) argument may be either purely numeric, such
' as "1.23" or in the form of a fractional value, such as "3/8".
'
' The function will return the decimal value of the evaluated string.
' An invalid numerical string will return an error message.
'
' ------------------------------------------------------------------------------
' Text_Part_Of(ArgString)
'
' Fetches only the symbol or text lable part from a single compound data
' argument string.  It is the counterpart of the Num_Part_Of() function above.
'
' ------------------------------------------------------------------------------
' ReSpaced(ArgString)
'
' Removes any extra spaces from within a compound element string.  This will
' replace all occurences of multiple spaces within a string with single spaces
' instead and also remove any leading or trailing spaces.
'
' ------------------------------------------------------------------------------
' DeSpaced(ArgString)
'
' Removes ALL spaces from within a compound element string.  This will push the
' characters together by removing any spaces that were between them and also
' remove any leading or trailing spaces.
'
' ------------------------------------------------------------------------------
' Error_In(Output_String)
'
' A Boolean function that checks for returned error messages from certain
' custom functions.  It returns Boolean "True" if an error message was found,
' otherwise it returns "False".
'
' Any NeoProgrammics custom error messages always start with "ERROR: " and
' this is what this function checks for in the string passed to it.
'
' ------------------------------------------------------------------------------
' Item(Index, FromString)
'
' This function returns an indexed item from within a delimited (FromString)
' containing several items.  Items are indexed from 1 upwards.
' It allows the fetching of the Nth delimited item within a string of items.
'
' ------------------------------------------------------------------------------
' Item_Count_In(DelimitedDataVector)
'
' This function counts the number of delimited items in a delimited data vector
' of the general format "Item1|Item2|Item3|Itemx ..."
'
' ------------------------------------------------------------------------------
' Occurrences_Of(SubStringA, WithinStringB)
'
' This function counts the number of occurrences of (SubStringA)
' in (WithinStringB).
'
' ------------------------------------------------------------------------------
' Delimited(DataString)
'
' This function delimits the variable contents of a string of data.  Numeric
' string data and text data are separated into delimited groups.  A series of
' numerical values separated by spaces may also be delimited.
'
' ------------------------------------------------------------------------------
' This function replace all occurrences of the substring in (InStringA) within
' the string (StringB) with the substring held in (WithStringC).
'
' Substitute(InStringA, StringB, ForStringC)
'
' ------------------------------------------------------------------------------
'

' ==============================================================================
' ==============================================================================
' ==============================================================================
'
' This function reads a string from the beginning (left end) and returns only
' the numerical characters.  It stops at the first non-numerical or space
' character encountered.  If there is no numeric part, then a null string is
' returned.
'
' This function can be used to parse a single string containing a numeric part
' followed by a text label or symbol and return the numeric part only.
'
' Since spaces are treated as non-numeric, if two numbers are separated by a
' space, only the first number will be returned and everything from the space
' onward will be ignored.
'
' The counterpart of this function is the Text_Part_Of() function which returns
' the label or symbol following the numeric part.
'
' NOTE: This function does NOT check for valid numbers - only if there are
' characters that can be part of a valid numeric string.
'
' Errors should be checked for by any routine(s) using the returned substring.
'
' Given the argument "12.34 mi", it will return the numeric string "12.3"
' Given the argument "12/34 ft", it will return the string "12/34"
'
' NOTE: The letter uppercase "E" is considered a numeric exponent marker, as
' in the value "1.5E-09".  This could cause problems for labels that begin
' with an uppercase "E".

'
' Spaces between the numeric and symbol parts are ignored.

  Public Static Function Num_Part_Of(ArgString)

  Dim Q As String
  Dim i As Integer

  Dim N As String
      N = "-+.0123456789/E"

  Q = Trim(ArgString)

' ----------------------------
' Check if pure numeric string

  If IsNumeric(Q) Then Num_Part_Of = DeSpaced(Q): Exit Function

' -----------------------
' Check for blank string

  If Q = "" Then Num_Part_Of = "": Exit Function

' --------------------------------------
' Find start of non-numeric part, if any

  For i = 1 To Len(Q)
      If InStr(N, Mid(Q, i, 1)) = 0 Then Exit For
  Next i

' -------------------------------------------
' Return null if single non-numeric character

  If i = 1 And InStr(N, Left(Q, 1)) = 0 Then _
     Num_Part_Of = "": Exit Function

  Num_Part_Of = DeSpaced(Left(Q, i - 1))

  End Function

' ==============================================================================
' ==============================================================================
' Function to evaluate a numerical value string given.  This function is a
' companion to the Num_Part_Of() and Text_Part_Of() functions.
'
' The Num_Part_Of() function extract a numerical value string from a single
' compound argument string, but does not evaluate the numerical value.
'
' This Num_Val_Of() function will evaluate a single numerical string given as
' either a purely numerical string or as a fractional value in the form "A/B".
'
' If there is no "/" character found, then it assumes the value is a plain
' numeric value and will return an error if the value is found to be invalid.

  Public Static Function Num_Val_Of(ArgString)
  
  Dim W As String
  Dim i As Integer
  Dim A As String
  Dim B As String

' Read numeric value from string argument.
  W = Num_Part_Of(ArgString)
  B = 1

'    Check for "/" character within string.
     i = InStr(1, W, "/")
  If i > 0 Then
     A = Left(W, i - 1)
     B = Trim(Mid(W, i + 1, Len(W)))

'    Report error if either A or B is non-numeric.
     If Not IsNumeric(A) Or Not IsNumeric(B) Then GoTo ERR_OUT

'    Check for division by zero error to prevent program crash
     A = Val(A): B = Val(B)
     If B = 0 Then Num_Val_Of = "ERROR: Division by zero in Num_Part_Of()" _
   : Exit Function

'    Return decimal value of A/B
     Num_Val_Of = A / B
     Exit Function

  End If

' Report error if non-fractional value is also non-numeric.
  If Not IsNumeric(W) Then GoTo ERR_OUT
 
' Return decimal value of A
  Num_Val_Of = Val(W)

  Exit Function

ERR_OUT:
  Num_Val_Of = "Num_Val_Of() ERROR: """ & ArgString & """ = Invalid argument."

  End Function

' ==============================================================================
' ==============================================================================

' This function returns only the symbol or text label following a numeric
' value.  It is the counterpart of the Num_Part_Of() function.
'
' Given the argument "12.34 mi", it will return the symblol "mi"
' Given the argument "12/34 km", it will return the symblol "km"
'
' Spaces between the numeric and symbol parts are ignored.
'
' NOTE: The letter uppercase "E" is considered a numeric exponent marker, as
' in the value "1.5E-09"

  Public Static Function Text_Part_Of(ArgString)

  Dim Q As String
  Dim C As String
  Dim i As Integer

  Dim N As String
      N = "-+.0123456789/E"

  Q = Trim(ArgString)

' -----------------------------------------------------
' Return null if pure numeric string or a null argument

  If IsNumeric(Q) Or Q = "" Then _
     Text_Part_Of = "": Exit Function

' ---------------------------------
' Find end of numeric part, if any

  For i = 1 To Len(Q)
      If InStr(N, Mid(Q, i, 1)) = 0 Then Exit For
  Next i

' -----------------------------------------------------------
' Return text or symbol characters following the numeric part

  Text_Part_Of = Trim(Mid(Q, i, Len(Q)))

  End Function

' ==============================================================================
' ==============================================================================
'
' Function to normalize the spacing between the elements of a string. If there
' is more than one space between the string elements, the extra spaces will will
' be removed leaving only one space between elements instead.
' Leading and trailing spaces are removed.  If there are no excess spaces, then
' no changes are made.

  Public Static Function ReSpaced(ArgString)

  Dim S  As String
  Dim i  As Integer
  
' Read string argument and trim off any leading and trailing spaces.
  S = Trim(ArgString)

' Replace any multiple spaces between elements with single spaces.
  While InStr(1, S, "  ") > 0
        i = InStr(1, S, "  ")
        S = Left(S, i - 1) & " " & Trim(Mid(S, i, Len(S)))
  Wend

  ReSpaced = S

  End Function

' ==============================================================================
' ==============================================================================
'
' Function to remove all spaces from within a string.  Leading and trailing
' and trailing spaces are also removed.  This will pack all the remaining
' characters together with no spaces in between elements or words.

  Public Static Function DeSpaced(ArgString)

  Dim W  As String

  Dim T  As String
      T = ""

  Dim S  As String
      S = ""

  Dim i  As Long
  
' Read string argument and trim off any leading and trailing spaces.
  S = Trim(ArgString)

'     Remove any spaces between string elements.
  For i = 1 To Len(S)
      W = Mid(S, i, 1): If W <> " " Then T = T & W
  Next i
  
  DeSpaced = T

  End Function

' ==============================================================================
' ==============================================================================
'
' This function returns the error status of the returned value of a function.
'
' If the returned string from a function contains the substring "ERROR", then
' it returns Boolean "True", otherwise "False".
'
' The NeoProgrammics modules use the convention of returning error messages
' beginning with "ERROR: " and followed by the message text.
'
' This makes it easier to detect if an error occurred within one of the custom
' functions.
'
' Just pass the returned string to this function as an argument to find out if
' an error was returned.
'
' This only applies to functions that are designed to detect and report errors.
' Some functions do not do this.
'
' This function is NOT case sensitive.

  Public Static Function Error_In(TestString) As Boolean

  If InStr(UCase(TestString), "ERROR") > 0 Then
     Error_In = True
  Else
     Error_In = False
  End If

  End Function

' ==============================================================================
' ==============================================================================
'
' This function gets any indexed item from within a delimited data vector.
' Items are numbered from 1 upwards.
'
' The delimiter character is the bar "|" or Chr$(124)
' If there is no delimiter, then (FromString) is returned as is.
' If item is not found, or index is invalid then a null string is returned
' rather than an error message.

  Public Static Function Item(Index, FromString)

  Dim i      As Long    ' Internal string pointer
  Dim Kount  As Long    ' Number of delimited items in (FromString)
  Dim N      As Long    ' Internal pointer index
  Dim W      As String  ' Work string
      W = ""

' Read input arguments
  N = Val(Index) - 1
  W = Trim(FromString)

' Count items in (FromString)
  Kount = Item_Count_In(W)

' Return a null string if N is out of proper range.
  If N < 0 Or N > Kount Then Item = "": Exit Function

' Return (FromString) unchanged if no delimiters were found (Kount=0).
  If Kount = 0 Then Item = Trim(W): Exit Function

' Return null string if N >= Kount
  If N >= Kount Then Item = "": Exit Function

' Find the (N)th item within (FromString)
  i = 0
  While i <= N - 1
        W = Mid(W, InStr(1, W, "|") + 1, Len(W))
        i = i + 1
  Wend

' Now, return the (N)th item found within (FromString).
' Any leading or trailing spaces attached to the extracted item are NOT cut off.
  i = InStr(1, W, "|")
  If i = 0 Then Item = W: Exit Function
  Item = Trim(Left(W, i - 1))

  End Function

' ==============================================================================
' ==============================================================================
' This function will return the counted number of items contained within a
' delimited data vector string.

  Public Static Function Item_Count_In(DelimitedDataVector)

  Dim i     As Integer ' Internal string pointer
  Dim Kount As Integer ' Item count
  Dim DVect As String  ' Copy of (DelimitedDataVector)

' Read data vector argument
  DVect = ReSpaced(Trim(DelimitedDataVector))
  DVect = Delimited(DVect)

' Remove any multiple delimiter occurences and replace by single delimiter.
  DVect = Substitute(DVect, "|", "||")

' Remove any delimiters at start of data vector
  While Left(DVect, 1) = "|"
        DVect = Trim(Mid(DVect, 2, Len(DVect)))
  Wend

' Remove any delimiters at end of data vector
  While Right(DVect, 1) = "|"
        DVect = Trim(Left(DVect, Len(DVect) - 1))
  Wend

' Return zero if data vector is empty of items
  If DVect = "" Then Item_Count_In = 0: Exit Function

' Count number of delimiter characters within data vector
  Kount = 0
  For i = 1 To Len(DVect)
          If Mid(DVect, i, 1) = "|" Then Kount = Kount + 1
  Next i

' Return number of delimited items counted
  Item_Count_In = Kount + 1

  End Function

' ==============================================================================
' ==============================================================================
' This function counts the number of occurrences of a given (SubStringA) within
' the string (WithinStringB).
'
' This function IS case sensitive.
'
' (SubStringA) is used exactly as given and may contain spaces or simply be all
' spaces.
'
' Any leading or trailing spaces of (WithinStringB) are ignored, however other
' spaces in (WithinStringB) WILL be retained.

  Public Static Function Occurrences_Of(SubStringA, WithinStringB)

  Dim StrA  As String   ' Exact (SubStringA) copy
  Dim StrB  As String   ' Trimmed (WithinStringB) copy

  Dim LenStrA  As Long  ' Length of unmodified (SubStringA)
  Dim P        As Long  ' Internal string pointer
  Dim Kount    As Long  ' Counted number of occurrences of (SubStringA)

' Read exact substring to be counted
  StrA = SubStringA

' Return zero if substring is null
  If StrA = "" Then Occurrences_Of = 0: Exit Function

' Get length of substring
  LenStrA = Len(StrA)

' Read string argument to search and trim off leading and trailing spaces
  StrB = Trim(WithinStringB)

' Count all occurrences of (StrA) within (StrB)
  Kount = 0
      P = 1
  Do Until P = 0
        P = InStr(1, StrB, StrA)
     If P > 0 Then _
     StrB = Mid(StrB, P + LenStrA, Len(StrB)): Kount = Kount + 1
  Loop

' Return result of substring count
  Occurrences_Of = Kount
  
  End Function

' ==============================================================================
' ==============================================================================
'
' This functions splits characters and numbers of (DataString) into delimited
' elements separated by the bar "|" or ANSI code 124 character.
' Leading and trailing spaces of (DataString) are ignored.
'
  Public Static Function Delimited(DataString)

  Dim U, W
 
  Dim NumChar As String
      NumChar = "+-0123456789./"

  Dim Arg      As String
  Dim i        As Integer

  Dim LastType As Integer
  Dim ThisType As Integer
      LastType = 1
      ThisType = 1

  Arg = ReSpaced(DataString)
    W = ""

  For i = 1 To Len(Arg)
      U = Mid(Arg, i, 1)
      If InStr(NumChar, U) = 0 Then ThisType = 0 Else ThisType = 1
      If ThisType <> LastType Then W = W & "|": LastType = ThisType
      W = W & U
  Next i

' Take care of slash problems
  While InStr(W, "|/") > 0
        i = InStr(W, "|/")
        W = Left(W, i - 1) & Mid(W, i + 1, Len(W))
  Wend
  While InStr(W, "/|") > 0
        i = InStr(W, "/|")
        W = Left(W, i) & Mid(W, i + 2, Len(W))
  Wend
  If Left(W, 1) = "|" Then W = Mid(W, 2, Len(W))

' Get rid of redundant spaces
  W = Substitute(W, "|", " |")
  W = Substitute(W, "|", "| ")
  If Left(W, 1) = "|" Then W = Mid(W, 2, Len(W))

' Remove dual delimiters
  W = Substitute(W, "|", "||")
  If Right(W, 1) = "|" Then W = Left(W, Len(W) - 1)
  
  Delimited = W

  End Function

' ==============================================================================
' ==============================================================================
'
' Substitute within (InStringA) the string (StringB) for (ForStringC).  This
' will replace ALL occurences of (ForStringC) with (StringB) within (InStringA).

' This function exists for compatiblity, because VB5 has no Replace() function
' like VB6 and this serves a similar purpose when needed.

  Public Static Function Substitute(InStringA, StringB, ForStringC)

  Dim StrA As String ' Main working string where replacement is to be made
  Dim StrB As String ' String that is to replace (ForStringC) within (InStringA)
  Dim StrC As String ' String that is to be replaced by (StringB)

  Dim i    As String ' Internal string pointer
  Dim LH   As String ' Left part of (InStringA) prior to substring (StringB)
  Dim RH   As String ' Right part of (InStringA) following substring (StringB)

  StrA = InStringA   ' Read value of string where replacement is to be made
  StrB = StringB     ' Read value of string that will replace (ForStringC)
  StrC = ForStringC  ' Read value of string that will be replaced by (StringB)

' Return result unchanged if either the first or final argument is null.
  If StrA = "" Or StrC = "" Then Substitute = InStringA: Exit Function

' Begin substitute replacement loop
  While InStr(StrA, StrC) > 0

' Split string in two halves on each side of substring to replace (StrC)
  i = InStr(StrA, StrC)
  LH = Left(StrA, i - 1): RH = Mid(StrA, i + Len(StrC), Len(StrA))
  
' Insert (StrB) in place of (StrC) within (StrA)
  StrA = LH & StrB & RH

' Repeat if not finished yet
  DoEvents
  Wend

' Return modified string result
  Substitute = StrA

  End Function

' ==============================================================================
' ==============================================================================
'

