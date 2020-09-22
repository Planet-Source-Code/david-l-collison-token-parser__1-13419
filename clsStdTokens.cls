VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStdTokens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------
' Class Name    : clsStdTokens
'
' Author        : David L. Collison - 12/04/2000
'
' Description   : This class will manipulate a text string and identify all
'                   tokens contained within the string.  Each token is seperated
'                   by a delimiter and may contain a token name and value.  Tokens
'                   can be retrieved using the position or name.
'
' Example       :
'   Bob,Sue,Sandy,Elaine - In this instance the comma is used to seperate each
'       name.  There are a total of four names or tokens identified.
'
'   name:bob;age:26;gender:m;height:6'1";weight:185 lbs;eyes:blue;hair:black
'       In this instance the semi-colon is used to seperate each token.  Each
'       token has a name and a value.
'
' Usage         :
'
'    Dim loTokens As New clsStdTokens
'    Dim liTokens As Integer
'
'    Debug.Print "Token Delimiter   : " & loTokens.sDelimiter
'
'    loTokens.sTokenList = "DataSource:tblAccounts;DataField:strName"
'    loTokens.gfTokenAdd "Locked:True"
'
'    Debug.Print "Max Tokens in list: " & Str(loTokens.iTokens)
'    Debug.Print "Full Token String : " & loTokens.sTokenList
'    Debug.Print "1st Token By Name : " & loTokens.gfTokenGet("DataSource")
'    Debug.Print "2nd Token By Name : " & loTokens.gfTokenGet("DataField")
'    Debug.Print "3rd Token By Name : " & loTokens.gfTokenGet("Locked")
'
'    For liTokens = 1 To loTokens.iTokens
'        Debug.Print "Token #" & Trim$(Str$(liTokens)); " By Index : " & loTokens.gfTokenGet(liTokens)
'    Next
'
'---------------------------------------------------------------------------------

Option Explicit

Private msTokenList As String           ' Tokens being parsed by the routine.
Private msDelimiter As String * 1       ' Delimiter used to identify tokens.

Private mcTokenList As New Collection   ' Tokens are stored in a collection.

Property Get iTokens() As Integer

    ' Allows the user to identify the total number of tokens in the list.
    
    iTokens = mcTokenList.Count
    
End Function

Property Get sDelimiter() As String

    ' Allows user to identify delimiter being used to parse the string.
    
    sDelimiter = msDelimiter

End Property

Property Let sDelimiter(ByVal psDelimiter As String)

    ' User can identify what they want to use as a delimiter.
    '   The ":" character is not allowed as it can be used to identify a tokens
    '   name vs. the token value.
    
    If (Len(psDelimiter) = 1) Then      ' Delimiter may only be 1 charachter.
        If (psDelimiter <> ":") Then    ' Don't allow the ":" character.
            msDelimiter = psDelimiter   ' Assign the delimiter.
        Else                            ' Display error message if ":" charachter.
            MsgBox "Delimiter assignment of ':' is invalid.  " & _
                "Used to identify keys.", _
                    vbExclamation, "Error - clsStdToken.sDelimiter"
        End If
    Else                                ' Display error if length > 1 charachter.
        MsgBox "Delimiter string may only contain 1 charachter.", _
            vbExclamation, "Error - clsStdToken.sDelimiter"
    End If
    
End Property

Property Get sTokenList() As String
    
    ' Allow the user to retrieve the token string.
    
    sTokenList = msTokenList
    
End Property

Property Let sTokenList(ByVal psTokenList As String)

    ' User sets the value of the token string being parsed.
    '   psTokenList - the string being passed in to the sTokenList property.

    Dim lsToken As String               ' Working string to store tokens.
    
    msTokenList = psTokenList           ' Assign to class variable.
    
    Set mcTokenList = New Collection    ' Reset the collection.
    
    Do While (Len(psTokenList) > 0)     ' Loop to retrieve tokens.
    
        If (InStr(1, psTokenList, msDelimiter) > 0) Then        ' Token found.
        
            ' Retrieve token.
            lsToken = Left$(psTokenList, InStr(1, psTokenList, msDelimiter) - 1)
            
            ' Reset token list.
            psTokenList = Right$(psTokenList, _
                Len(psTokenList) - InStr(1, psTokenList, msDelimiter))
            
        Else                                                    ' No more tokens.
        
            ' Retrieve last token.
            lsToken = psTokenList
            
            ' Set token list to nothing.
            psTokenList = Space$(0)
            
        End If
        
        ' Assign token to collection.
        pfTokenAdd lsToken
    
    Loop
    
End Property

Private Sub Class_Initialize()
    ' Set initial delimiter to be the semi-colon.
    msDelimiter = ";"
End Sub

Private Function pfTokenAdd(ByVal psToken As String)

    ' Add a token to an existing list of tokens.

    Dim msKey As String                         ' Temporary token name.

    If (InStr(1, psToken, ":") > 0) Then        ' Identify if token has name.
        
        ' Retrieve name of the token.
        msKey = Left$(psToken, InStr(1, psToken, ":") - 1)
        ' Retrieve the token.
        psToken = Right$(psToken, Len(psToken) - InStr(1, psToken, ":"))
        ' Add token to the collection.
        mcTokenList.Add psToken, msKey
    Else                                        ' If token has no name.
        ' Add token to the collection.
        mcTokenList.Add psToken
    End If

End Function

Public Function gfTokenAdd(ByVal psToken As String) As Boolean

    ' Allows user to add a token to the collection of current tokens.

    pfTokenAdd psToken      ' Add to the collection.
    
    ' Add token to current string.
    msTokenList = msTokenList + msDelimiter + psToken

End Function

Public Function gfTokenExists(ByVal psToken As String, _
    Optional ByVal pbIgnoreCase As Boolean) As Boolean

    ' Allows the user to search for a token.

    Dim lvToken As Variant                      ' Token being examined.
    
    If (IsMissing(pbIgnoreCase)) Then           ' Do we ignore case during search?
        pbIgnoreCase = False
    End If
    
    gfTokenExists = False                       ' No token found yet.

    For Each lvToken In mcTokenList             ' Parse each token in collection.
    
        If (pbIgnoreCase) Then                  ' Ignore the case when examining.
            If (UCase$(lvToken) = UCase$(psToken)) Then
                gfTokenExists = True            ' Token found.
                Exit For                        ' Exit loop
            End If
        Else                                    ' We do not ignore the case.
            If (lvToken = psToken) Then
                gfTokenExists = True            ' Token found.
                Exit For                        ' Exit loop.
            End If
        End If
        
    Next
    
End Function

Public Function gfTokenGet(Optional pvIndex As Variant) As String

    ' Allow the user to retrieve a Token by either index position or name.

    Dim liLoops As Integer
    
    If (IsMissing(pvIndex)) Then
        pvIndex = 1
    End If
    
    Select Case VarType(pvIndex)
    Case Is = vbEmpty Or vbNull Or vbObject Or vbError Or vbBoolean _
        Or vbVariant Or vbDataObject Or vbUserDefinedType Or vbArray
        
        pvIndex = 1                                 ' Assign default value.
    
    Case Is = vbInteger Or vbString
        ' Do nothing, leave as integer or string.
        
    Case Is = vbLong Or vbSingle Or vbDouble Or vbCurrency Or vbDate _
        Or vbDecimal Or vbByte
        
        pvIndex = CInt(pvIndex)                     ' Force index to integer.
    
        If (pvIndex < 1) Then                       ' Validate token within index.
            pvIndex = 1
        ElseIf (pvIndex > mcTokenList.Count) Then
            pvIndex = mcTokenList.Count
        End If
    
    End Select

    gfTokenGet = mcTokenList(pvIndex)               ' Retrieve token from list.

End Function
