Attribute VB_Name = "modAccounts"
Option Explicit

Public Const MaxCharPerAccount As Byte = 20
Const KeyTokenAuth As String = "demo"


' Token

Public Function ValidJWT(ByVal sToken As String) As Boolean
'***************************************************
'Autor: ^[GS]^
'Last Modification: 29/01/2022 - ^[GS]^
'
'***************************************************

    If Len(sToken) < 32 Then
        ValidJWT = False
        Exit Function
    End If
    
    Dim sItems() As String
    sItems = Split(sToken, ".")
    ValidJWT = False

    If (UBound(sItems) + 1) = 3 Then
        If Len(sItems(0)) = 36 And Len(sItems(1)) > 180 And Len(sItems(2)) = 43 Then
            ValidJWT = True
        End If
    End If

End Function

' Accounts

Private Function GetHTMLSource(ByVal sURL As String) As String
'***************************************************
'Autor: ^[GS]^
'Last Modification: 29/01/2022 - ^[GS]^
'
'***************************************************

    Dim xmlHttp As Object

    Set xmlHttp = CreateObject("MSXML2.XmlHttp")
    xmlHttp.Open "GET", sURL, False
    xmlHttp.send
    GetHTMLSource = xmlHttp.responseText
    Set xmlHttp = Nothing
    
End Function


Public Function ConnectAccount(ByVal UserIndex As Integer, ByVal Token As String) As Boolean
'***************************************************
'Autor: ^[GS]^
'Last Modification: 29/01/2022 - ^[GS]^
'
'***************************************************

    ConnectAccount = False ' Por defecto, es FALSE
    
    With UserList(UserIndex)
    
        If .flags.AccountLogged Then ' Ya está logueado
            Call LogCheating("La cuenta " & .AccountName & " ha intentado loguear desde la IP " & .ip)
            Call CloseSocketSL(UserIndex)
            Call Cerrar_Usuario(UserIndex)
            Exit Function
        End If
        
        Dim sURL As String
        Dim sGet As String
        
        sURL = "https://www.gs-zone.org/login_auth.php?&code=gszoneao&key=" & KeyTokenAuth & "&token=" & Token
        sGet = GetHTMLSource(sURL)
        Dim sItems() As String
    
        sItems = Split(sGet, ";")
        If (UBound(sItems) + 1) <> 3 Then
            Call WriteErrorMsg(UserIndex, "Token invalido. Vuelve a acceder para obtener un nuevo token.")
            Call FlushBuffer(UserIndex)
            Call CloseSocket(UserIndex)
            Exit Function
        End If
        
        .AccountID = CLng(sItems(0))
        .AccountName = sItems(1)
        .AccountHash = sItems(2)
        
        If Not FileExist(pathAccounts & .AccountHash & ".ach", vbNormal) Then
            Call WriteVar(pathAccounts & .AccountHash & ".ach", "INIT", "ID", .AccountID)
            Call WriteVar(pathAccounts & .AccountHash & ".ach", "INIT", "Account", .AccountName)
            Call WriteVar(pathAccounts & .AccountHash & ".ach", "INIT", "LastIP", .ip)
        End If
        
        .flags.AccountLogged = True
        
        ConnectAccount = True
        
    End With

End Function


Public Function LoadAccount(ByVal UserIndex As Integer)
'***************************************************
'Autor: ^[GS]^
'Last Modification: 29/01/2022 - ^[GS]^
'
'***************************************************

On Error GoTo ErrorHandler

    Dim AccountName As String
    Dim AccountHash As String
    
    AccountName = UserList(UserIndex).AccountName
    AccountHash = UserList(UserIndex).AccountHash

    Dim account            As clsIniManager
    Dim CharFile           As clsIniManager
    Dim i                  As Long
    Dim NumberOfCharacters As Byte
    Dim Characters()       As AccountUser
    Dim CurrentCharacter   As String

    Set account = New clsIniManager
    Set CharFile = New clsIniManager
    
    Call WriteVar(pathAccounts & AccountHash & ".ach", "INIT", "LastIP", UserList(UserIndex).ip)

    Call account.Initialize(pathAccounts & AccountHash & ".ach")
    NumberOfCharacters = val(account.GetValue("INIT", "CantidadPersonajes"))

    If NumberOfCharacters > 0 Then
        ReDim Characters(1 To NumberOfCharacters) As AccountUser

        For i = 1 To NumberOfCharacters
            CurrentCharacter = account.GetValue("PERSONAJES", "Personaje" & i)
            
            If Len(CurrentCharacter) > 0 Then
                Call CharFile.Initialize(pathChars & CurrentCharacter & ".chr")
                Characters(i).Name = CurrentCharacter
                Characters(i).Body = val(CharFile.GetValue("INIT", "Body"))
                Characters(i).Head = val(CharFile.GetValue("INIT", "Head"))
                Characters(i).Weapon = val(CharFile.GetValue("INIT", "Arma"))
                Characters(i).Shield = val(CharFile.GetValue("INIT", "Escudo"))
                Characters(i).Helmet = val(CharFile.GetValue("INIT", "Casco"))
                Characters(i).Class = val(CharFile.GetValue("INIT", "Clase"))
                Characters(i).Race = val(CharFile.GetValue("INIT", "Raza"))
                Characters(i).Map = CharFile.GetValue("INIT", "Position")
                Characters(i).Level = val(CharFile.GetValue("STATS", "ELV"))
                Characters(i).Gold = val(CharFile.GetValue("STATS", "GLD"))
                Characters(i).LastConnect = LTrim$(ReadField(2, CharFile.GetValue("INIT", "LASTIP1"), 45))
                Characters(i).Criminal = (val(CharFile.GetValue("REP", "Promedio")) < 0)
                Characters(i).Dead = CBool(val(CharFile.GetValue("FLAGS", "Muerto")))
                Characters(i).GameMaster = EsGmChar(CurrentCharacter)
            End If
        Next i

    End If

    Set account = Nothing
    Set CharFile = Nothing

    Call modProtocol.WriteAccountChars(UserIndex, NumberOfCharacters, Characters)

    Exit Function
ErrorHandler:
    Call LogError("Error in LoginAccountToken: " & AccountHash & ". " & Err.Number & " - " & Err.description)


End Function

Function NumberOfCharacters(AccountHash As String) As Integer
'***************************************************
'Autor: ^[GS]^
'Last Modification: 26/02/2022 - ^[GS]^
'
'***************************************************

    If Not FileExist(pathAccounts & AccountHash & ".ach", vbNormal) Then
        NumberOfCharacters = 0
        Exit Function
    End If

    Dim account As clsIniManager
    Set account = New clsIniManager
    Call account.Initialize(pathAccounts & AccountHash & ".ach")
    NumberOfCharacters = val(account.GetValue("INIT", "CantidadPersonajes"))

End Function

Function IsCharInAccount(ByVal UserIndex As Integer, ByVal charName As String) As Boolean
'***************************************************
'Autor: ^[GS]^
'Last Modification: 01/03/2022 - ^[GS]^
'
'***************************************************

    If Len(UserList(UserIndex).AccountHash) = 0 Then
        Exit Function
    End If

    Dim NumChars As Byte
    Dim account As clsIniManager
    Set account = New clsIniManager
    Call account.Initialize(pathAccounts & UserList(UserIndex).AccountHash & ".ach")
    NumChars = val(account.GetValue("INIT", "CantidadPersonajes"))
    If NumChars > 0 Then
        Dim LoopC As Byte
        Dim Name As String
        For LoopC = 1 To NumChars
            Name = UCase$(account.GetValue("PERSONAJES", "Personaje" & LoopC))
            If Name = UCase$(charName) Then
                IsCharInAccount = True
                Exit Function
            End If
        Next
    End If
    IsCharInAccount = False

End Function

Function AddCharacterToAccount(ByVal UserIndex As Integer, ByVal charName As String) As Boolean
'***************************************************
'Autor: ^[GS]^
'Last Modification: 26/02/2022 - ^[GS]^
'
'***************************************************

    If Len(charName) = 0 Then
        Exit Function
    End If

    If Not FileExist(pathAccounts & UserList(UserIndex).AccountHash & ".ach", vbNormal) Then
        Exit Function
    End If

    Dim account As clsIniManager
    Set account = New clsIniManager
    Call account.Initialize(pathAccounts & UserList(UserIndex).AccountHash & ".ach")
    
    Dim NumberOfCharacters As Byte
    NumberOfCharacters = val(account.GetValue("INIT", "CantidadPersonajes"))
    NumberOfCharacters = NumberOfCharacters + 1
    
    If NumberOfCharacters > MaxCharPerAccount Then
        Exit Function
    End If
    
    Call WriteVar(pathAccounts & UserList(UserIndex).AccountHash & ".ach", "INIT", "CantidadPersonajes", NumberOfCharacters)
    Call WriteVar(pathAccounts & UserList(UserIndex).AccountHash & ".ach", "PERSONAJES", "Personaje" & NumberOfCharacters, charName)
    
    AddCharacterToAccount = True

End Function

Sub DeleteCharacterFromAccount(ByVal UserIndex As Integer, ByVal charName As String)
'***************************************************
'Autor: ^[GS]^
'Last Modification: 01/03/2022 - ^[GS]^
'
'***************************************************

    If Len(charName) = 0 Then
        Exit Sub
    End If

    If Not FileExist(pathAccounts & UserList(UserIndex).AccountHash & ".ach", vbNormal) Then
        Exit Sub
    End If

    Dim account As clsIniManager
    Set account = New clsIniManager
    Call account.Initialize(pathAccounts & UserList(UserIndex).AccountHash & ".ach")
    
    Dim NumberOfCharacters As Byte
    NumberOfCharacters = val(account.GetValue("INIT", "CantidadPersonajes"))
    
    If NumberOfCharacters = 0 Then
        Exit Sub
    End If
    
    Dim LoopC As Byte
    Dim CurrentCharacter As String
    Dim Founded As Byte
    
    For LoopC = 1 To NumberOfCharacters
        CurrentCharacter = account.GetValue("PERSONAJES", "Personaje" & LoopC)
        If UCase$(CurrentCharacter) = UCase$(charName) Then
            Founded = LoopC
        Else
            If Founded > 0 Then
                Call WriteVar(pathAccounts & UserList(UserIndex).AccountHash & ".ach", "PERSONAJES", "Personaje" & LoopC - 1, CurrentCharacter)
            End If
        End If
    Next
    If Founded > 0 Then
        Call WriteVar(pathAccounts & UserList(UserIndex).AccountHash & ".ach", "PERSONAJES", "Personaje" & NumberOfCharacters, "")
    End If
    
    NumberOfCharacters = NumberOfCharacters - 1
    Call WriteVar(pathAccounts & UserList(UserIndex).AccountHash & ".ach", "INIT", "CantidadPersonajes", NumberOfCharacters)
    
    If FileExist(pathChars & UCase$(charName) & ".chr", vbNormal) Then
        Kill pathChars & UCase$(charName) & ".chr"
    End If

End Sub
