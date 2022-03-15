Attribute VB_Name = "modMySQL"
#If Mysql = 1 Then

'************************************************************************
'************************************************************************
'************************************************************************
'*****Descripcion: Base de Datos Mysql Argentum Online V. 0.13***********
'*****Autor: Jose Ignacio Castelli ( Fedudok )***************************
'*****Fecha: 21/7/2011***************************************************
'************************************************************************
'************************************************************************
'************************************************************************

Option Explicit
Public Con As ADODB.Connection
Public Const mySQLHost As String = "localhost" ' host de DB
Public Const mySQLBase As String = "gszao" ' tabla de DB
Public Const mySQLUser As String = "root" ' usuario de DB
Public Const mySQLPass As String = "123456" ' contraseña de DB

Public Sub CargarDB()
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************

On Error GoTo ErrHandler

    Set Con = New ADODB.Connection
    Con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" & _
                "SERVER=" & mySQLHost & "; " & _
                "DATABASE=" & mySQLBase & ";" & _
                "UID=" & mySQLUser & ";" & _
                "PWD=" & mySQLPass & "; OPTION=3"
    
    Con.CursorLocation = adUseClient
    Con.Open
    Exit Sub
    
ErrHandler:
    MsgBox Err.description
    End
End Sub

Public Function ChangeBan(ByVal name As String, ByVal Baneado As Byte) As Boolean
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
    Dim Orden As String
    Dim RS As New ADODB.Recordset
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(name) & "'")
        If RS.BOF Or RS.EOF Then Exit Function
        
        Orden = "UPDATE `charflags` SET"
        Orden = Orden & " IndexPJ=" & RS!IndexPJ
        Orden = Orden & ",Nombre='" & UCase$(name) & "'"
        Orden = Orden & ",Ban=" & Baneado
        Orden = Orden & " WHERE IndexPJ=" & RS!IndexPJ

        Call Con.Execute(Orden)
    Set RS = Nothing

End Function


Public Sub CerrarDB()
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Error GoTo ErrHandle
    Con.Close
    Set Con = Nothing
    Exit Sub
ErrHandle:
    Call LogError("CerrarDB " & Err.description & " " & Err.Number)
    End
    
End Sub
Public Sub SaveUserSQL(userIndex As Integer, Optional insertPj As Boolean = False)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
        
    Dim iPJ As Integer
       
    If insertPj Then
        iPJ = Insert_New_Table(UserList(userIndex).name)
    Else
        iPJ = GetIndexPJ(UserList(userIndex).name)
    End If

    SaveUserFlags userIndex, iPJ
    SaveUserStats userIndex, iPJ
    SaveReputacion userIndex, iPJ
    SaveUserInit userIndex, iPJ
    SaveUserInv userIndex, iPJ
    SaveUserBank userIndex, iPJ
    SaveUserHechi userIndex, iPJ
    SaveUserAtrib userIndex, iPJ
    SaveUserSkill userIndex, iPJ
    SaveUserFaccion userIndex, iPJ
    
    Exit Sub

End Sub

Sub SaveUserHechi(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charhechizos` SET"
    str = str & " IndexPJ=" & iPJ
    For i = 1 To MAXUSERHECHIZOS
        str = str & ",H" & i & "=" & mUser.Stats.UserHechizos(i)
    Next i
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    
    Exit Sub
ErrHandle:
    Resume Next
End Sub




Sub SaveReputacion(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
        '************************************************************************
    Set RS = New ADODB.Recordset
    
    Set RS = Con.Execute("SELECT * FROM `reputacion` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    Dim Pena As Integer
    
    Set RS = Con.Execute("SELECT * FROM `reputacion` WHERE IndexPJ=" & iPJ)
    str = "UPDATE `reputacion` SET"
    str = str & " IndexPJ=" & iPJ
    str = str & ",Asesino=" & mUser.Reputacion.AsesinoRep
    str = str & ",Bandido=" & mUser.Reputacion.BandidoRep
    str = str & ",Burguesia=" & mUser.Reputacion.BurguesRep
    str = str & ",Ladrones=" & mUser.Reputacion.LadronesRep
    str = str & ",Nobles=" & mUser.Reputacion.NobleRep
    str = str & ",Plebe=" & mUser.Reputacion.PlebeRep
    str = str & ",Promedio=" & mUser.Reputacion.Promedio
    
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    'Grabamos Estados
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub





Sub SaveUserFlags(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 08/06/2012 - ^[GS]^
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
        '************************************************************************
    Set RS = New ADODB.Recordset
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & mUser.name & "'")
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    Dim Pena As Integer
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & iPJ)
    str = "UPDATE `charflags` SET"
    str = str & " IndexPJ=" & iPJ
    str = str & ",Nombre='" & mUser.name & "'"
    str = str & ",Ban=" & mUser.flags.Ban
    str = str & ",Navegando=" & mUser.flags.Navegando
    str = str & ",Envenenado=" & mUser.flags.Envenenado
    str = str & ",Pena=" & Pena * 60
    str = str & ",Paralizado=" & mUser.flags.Paralizado
    str = str & ",Desnudo=" & mUser.flags.Desnudo
    str = str & ",Sed=" & mUser.flags.Sed
    str = str & ",Hambre=" & mUser.flags.Hambre
    str = str & ",Escondido=" & mUser.flags.Escondido
    str = str & ",Muerto=" & mUser.flags.Muerto
    str = str & ",LastMap=" & mUser.flags.lastMap
    str = str & ",SkillsAsignados=" & mUser.Counters.AsignedSkills
    str = str & ",NPCSMUERTES=" & mUser.Stats.NPCsMuertos
    str = str & ",USERMUERTES=" & mUser.Stats.UsuariosMatados
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    'Grabamos Estados
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub

Sub SaveUserFaccion(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charfaccion` SET"
    
    'Graba Faccion
    str = str & " IndexPJ=" & iPJ
    str = str & ",EjercitoReal=" & mUser.fAccion.ArmadaReal
    str = str & ",EjercitoCaos=" & mUser.fAccion.FuerzasCaos
    str = str & ",CiudMatados=" & mUser.fAccion.CiudadanosMatados
    str = str & ",CaosMatados=" & mUser.fAccion.CriminalesMatados
    str = str & ",FechaIngreso='" & mUser.fAccion.FechaIngreso & "'"
    str = str & ",MatadosIngreso=" & mUser.fAccion.MatadosIngreso
    str = str & ",NextRecompenza=" & mUser.fAccion.NextRecompensa
    str = str & ",NivelIngreso=" & mUser.fAccion.NivelIngreso
    str = str & ",RecibioArmaduraCaos=" & mUser.fAccion.RecibioArmaduraCaos
    str = str & ",RecibioArmaduraReal=" & mUser.fAccion.RecibioArmaduraReal
    str = str & ",RecompensasCaos=" & mUser.fAccion.RecompensasCaos
    str = str & ",RecompensasReal=" & mUser.fAccion.RecompensasReal
    str = str & ",Reenlistadas=" & mUser.fAccion.Reenlistadas
    str = str & ",RecibioExpInicialCaos=" & mUser.fAccion.RecibioExpInicialCaos
    str = str & ",RecibioExpInicialReal=" & mUser.fAccion.RecibioExpInicialReal

    
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserInit(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charinit` SET"
    str = str & " IndexPJ=" & iPJ
    str = str & ",Genero=" & mUser.Genero
    str = str & ",Raza=" & mUser.raza
    str = str & ",Hogar=" & mUser.Hogar
    str = str & ",Clase=" & mUser.clase
    str = str & ",Heading=" & mUser.Char.heading
    str = str & ",Head=" & mUser.OrigChar.Head
    str = str & ",Body=" & mUser.Char.Body
    str = str & ",Arma=" & mUser.Char.WeaponAnim
    str = str & ",Escudo=" & mUser.Char.ShieldAnim
    str = str & ",Casco=" & mUser.Char.CascoAnim
    str = str & ",LastIP='" & mUser.ip & "'"
    str = str & ",Mapa=" & mUser.Pos.Map
    str = str & ",X=" & mUser.Pos.X
    str = str & ",Y=" & mUser.Pos.Y
    str = str & ",Desc='" & mUser.desc & "'"
    str = str & ",Password='" & mUser.Password & "'"
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub


Sub SaveUserPosition(ByVal iPJ As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'************************************************************************
'Autor: ^[GS]^
'Fecha: 31/05/2012
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    If InMapBounds(Map, X, Y) = False Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charinit` SET"
    str = str & " Mapa=" & Map
    str = str & ",X=" & X
    str = str & ",Y=" & Y
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserInv(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charinvent` SET"
    str = str & " IndexPJ=" & iPJ
    For i = 1 To MAX_INVENTORY_SLOTS
        str = str & ",OBJ" & i & "=" & mUser.Invent.Object(i).ObjIndex
        str = str & ",CANT" & i & "=" & mUser.Invent.Object(i).Amount
    Next i
    str = str & ",CASCOSLOT=" & mUser.Invent.CascoEqpSlot
    str = str & ",ARMORSLOT=" & mUser.Invent.ArmourEqpSlot
    str = str & ",SHIELDSLOT=" & mUser.Invent.EscudoEqpSlot
    str = str & ",WEAPONSLOT=" & mUser.Invent.WeaponEqpSlot
    str = str & ",MUNICIONSLOT=" & mUser.Invent.MunicionEqpSlot
    str = str & ",BARCOSLOT=" & mUser.Invent.BarcoSlot
    str = str & ",ANILLOSLOT=" & mUser.Invent.AnilloEqpSlot
    str = str & ",MOCHILASLOT=" & mUser.Invent.MochilaEqpSlot
    str = str & " WHERE IndexPJ=" & iPJ
    
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
    
ErrHandle:
    Resume Next
    
End Sub
Sub SaveUserBank(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charbanco` SET"
    str = str & " IndexPJ=" & iPJ
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        str = str & ",OBJ" & i & "=" & mUser.BancoInvent.Object(i).ObjIndex
        str = str & ",CANT" & i & "=" & mUser.BancoInvent.Object(i).Amount
    Next i
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserStats(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
     
    str = "UPDATE `charstats` SET"
    str = str & " IndexPJ=" & iPJ
    str = str & ",GLD=" & mUser.Stats.GLD
    str = str & ",BANCO=" & mUser.Stats.Banco
    str = str & ",MaxHP=" & mUser.Stats.MaxHp
    str = str & ",MinHP=" & mUser.Stats.MinHp
    str = str & ",MaxMAN=" & mUser.Stats.MaxMAN
    str = str & ",MinMAN=" & mUser.Stats.MinMAN
    str = str & ",MinSTA=" & mUser.Stats.MinSta
    str = str & ",MaxSTA=" & mUser.Stats.MaxSta
    str = str & ",MaxHIT=" & mUser.Stats.MaxHIT
    str = str & ",MinHIT=" & mUser.Stats.MinHIT
    str = str & ",MaxAGU=" & mUser.Stats.MaxAGU
    str = str & ",MinAGU=" & mUser.Stats.MinAGU
    str = str & ",MaxHAM=" & mUser.Stats.MaxHam
    str = str & ",MinHAM=" & mUser.Stats.MinHam
    str = str & ",SkillPtsLibres=" & mUser.Stats.SkillPts
    str = str & ",Exp=" & mUser.Stats.Exp
    str = str & ",ELV=" & mUser.Stats.ELV
    str = str & ",NpcsMuertes=" & mUser.Stats.NPCsMuertos
    str = str & ",ELU=" & mUser.Stats.ELU
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserAtrib(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charatrib` SET"
    str = str & " IndexPJ=" & iPJ
    For i = 1 To NUMATRIBUTOS
        str = str & ",AT" & i & "=" & mUser.Stats.UserAtributosBackUP(i)
    Next i
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Sub SaveUserSkill(ByVal userIndex As Integer, ByVal iPJ As Integer)
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Local Error GoTo ErrHandle
    Dim RS As ADODB.Recordset
    Dim mUser As User
    Dim i As Byte
    Dim str As String
    
    mUser = UserList(userIndex)
    
    If Len(mUser.name) = 0 Then Exit Sub
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then Exit Sub
    Set RS = Nothing
    
    str = "UPDATE `charskills` SET"
    str = str & " IndexPJ=" & iPJ
    
    For i = 1 To NUMSKILLS
        str = str & ",SK" & i & "=" & mUser.Stats.UserSkills(i)
        str = str & ",ELUSK" & i & "=" & mUser.Stats.EluSkills(i)
        str = str & ",EXPSK" & i & "=" & mUser.Stats.ExpSkills(i)
    Next i
    
    

    
    str = str & " WHERE IndexPJ=" & iPJ
    Call Con.Execute(str)
    '************************************************************************
    Exit Sub
ErrHandle:
    Resume Next
End Sub
Function LoadUserSQL(userIndex As Integer, ByVal name As String) As Boolean
On Error GoTo ErrHandler
Dim i As Integer
Dim RS As New ADODB.Recordset
Dim iPJ  As Integer

With UserList(userIndex)

    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & name & "'")
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
    
        iPJ = RS!IndexPJ
    Set RS = Nothing
    '************************************************************************
    
    .IndexPJ = iPJ
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE IndexPJ=" & iPJ)
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If

    .flags.Ban = RS!Ban
    .flags.Navegando = RS!Navegando
    .flags.Envenenado = RS!Envenenado
    .Counters.Pena = RS!Pena * 60
    .flags.Paralizado = RS!Paralizado
    .flags.Desnudo = RS!Desnudo
    .flags.Sed = RS!Sed
    .flags.Hambre = RS!Hambre
    .flags.Escondido = RS!Escondido
    .flags.Muerto = RS!Muerto

    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charfaccion` WHERE IndexPJ=" & iPJ)
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    ' Carga Faccion
    .fAccion.ArmadaReal = RS!EjercitoReal
    .fAccion.FuerzasCaos = RS!EjercitoCaos
    .fAccion.CiudadanosMatados = RS!CiudMatados
    ' Fin Carga Faccion
    
    Set RS = Nothing
    '************************************************************************

    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charatrib` WHERE IndexPJ=" & iPJ)
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    For i = 1 To NUMATRIBUTOS
        .Stats.UserAtributos(i) = RS.Fields("AT" & i)
        .Stats.UserAtributosBackUP(i) = .Stats.UserAtributos(i)
    Next i
    
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charguild` WHERE IndexPJ=" & iPJ)
        If RS.BOF Or RS.EOF Then
            LoadUserSQL = False
            Exit Function
        End If
        
        UserList(userIndex).GuildIndex = RS!GuildIndex
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charskills` WHERE IndexPJ=" & iPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To NUMSKILLS
        .Stats.UserSkills(i) = RS.Fields("SK" & i)
    Next i
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charbanco` WHERE IndexPJ=" & iPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        .BancoInvent.Object(i).ObjIndex = RS.Fields("OBJ" & i)
        .BancoInvent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinvent` WHERE IndexPJ=" & iPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAX_INVENTORY_SLOTS
        .Invent.Object(i).ObjIndex = RS.Fields("OBJ" & i)
        .Invent.Object(i).Amount = RS.Fields("CANT" & i)
    Next i
    .Invent.CascoEqpSlot = RS!CASCOSLOT
    .Invent.ArmourEqpSlot = RS!ARMORSLOT
    .Invent.EscudoEqpSlot = RS!SHIELDSLOT
    .Invent.WeaponEqpSlot = RS!WeaponSlot
    .Invent.MunicionEqpSlot = RS!MunicionSlot
    .Invent.BarcoSlot = RS!BarcoSlot
    
    Set RS = Nothing

    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charhechizos` WHERE IndexPJ=" & iPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    For i = 1 To MAXUSERHECHIZOS
        .Stats.UserHechizos(i) = RS.Fields("H" & i)
    Next i
    
    Set RS = Nothing
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charstats` WHERE IndexPJ=" & iPJ)
    
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    .Stats.GLD = RS!GLD
    .Stats.Banco = RS!Banco
    .Stats.MaxHp = RS!MaxHp
    .Stats.MinHp = RS!MinHp
    .Stats.MinSta = RS!MinSta
    .Stats.MaxSta = RS!MaxSta
    .Stats.MaxMAN = RS!MaxMAN
    .Stats.MinMAN = RS!MinMAN
    .Stats.MaxHIT = RS!MaxHIT
    .Stats.MinHIT = RS!MinHIT
    .Stats.MinAGU = RS!MinAGU
    .Stats.MinHam = RS!MinHam
    .Stats.MaxAGU = 100
    .Stats.MaxHam = 100
    .Stats.SkillPts = RS!SkillPtsLibres
    .Stats.Exp = RS!Exp
    .Stats.ELV = RS!ELV
    .Stats.NPCsMuertos = RS!NpcsMuertes
    .Stats.ELU = RS!ELU
    
    Set RS = Nothing
    
    If .Stats.MinAGU < 1 Then .flags.Sed = 1
    If .Stats.MinHam < 1 Then .flags.Hambre = 1
    If .Stats.MinHp < 1 Then .flags.Muerto = 1
    
    '************************************************************************
    
    '************************************************************************
    Set RS = Con.Execute("SELECT * FROM `charinit` WHERE IndexPJ=" & iPJ)
    If RS.BOF Or RS.EOF Then
        LoadUserSQL = False
        Exit Function
    End If
    
    .Genero = RS!Genero
    .raza = RS!raza
    .Hogar = RS!Hogar
    .clase = RS!clase
    .Char.heading = RS!heading
    .OrigChar.Head = RS!Head
    .Char.Body = RS!Body
    .Char.WeaponAnim = RS!Arma
    .Char.ShieldAnim = RS!Escudo
    .Char.CascoAnim = RS!casco
    .ip = RS!lastip
    .Pos.Map = RS!mapa
    .Pos.X = RS!X
    .Pos.Y = RS!Y
    
    If .flags.Muerto = 0 Then
        .Char = .OrigChar
        Call VerObjetosEquipados(userIndex)
    Else
        .Char.Body = iCuerpoMuerto
        .Char.Head = iCabezaMuerto
        .Char.WeaponAnim = NingunArma
        .Char.ShieldAnim = NingunEscudo
        .Char.CascoAnim = NingunCasco
    End If
    
    Set RS = Nothing


    '************************************************************************
    
    '************************************************************************
    
    
    LoadUserSQL = True
    
    If Len(.desc) >= 80 Then .desc = Left$(.desc, 80)

    .Stats.MaxAGU = 100
    .Stats.MaxHam = 100

End With

Exit Function

ErrHandler:
    Call LogError("Error en LoadUserSQL. N:" & name & " - " & Err.Number & "-" & Err.description)
    Set RS = Nothing
    
End Function


Public Function BANCheckDB(ByVal name As String) As Boolean
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
    Dim RS As New ADODB.Recordset
    Dim Baneado As Byte
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(name) & "'")
        If RS.BOF Or RS.EOF Then Exit Function
    
        Baneado = RS!Ban
        BANCheckDB = (Baneado = 1)
    Set RS = Nothing

End Function

Function ExistePersonaje(name As String) As Boolean
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
    Dim RS As New ADODB.Recordset
    
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(name) & "'")
        If RS.BOF Or RS.EOF Then Exit Function
    Set RS = Nothing
    
    ExistePersonaje = True
End Function

Public Function GetIndexPJ(ByVal name As String) As Integer
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Error GoTo Err
    Dim RS As New ADODB.Recordset
    Dim IndexPJ As Long

    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & UCase$(name) & "'")
        If RS.BOF Or RS.EOF Then
            GoTo Err
        Else
            GetIndexPJ = RS!IndexPJ
        End If
    Set RS = Nothing
    Exit Function
    
Err:
    Set RS = Nothing
    GetIndexPJ = 0
    Exit Function
End Function

Public Sub VerObjetosEquipados(userIndex As Integer)

'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************

With UserList(userIndex).Invent
    If .CascoEqpSlot Then
        .Object(.CascoEqpSlot).Equipped = 1
        .CascoEqpObjIndex = .Object(.CascoEqpSlot).ObjIndex
        UserList(userIndex).Char.CascoAnim = ObjData(.CascoEqpObjIndex).CascoAnim
    Else
        UserList(userIndex).Char.CascoAnim = NingunCasco
    End If
    
    If .BarcoSlot Then .BarcoObjIndex = .Object(.BarcoSlot).ObjIndex
    
    If .ArmourEqpSlot Then
        .Object(.ArmourEqpSlot).Equipped = 1
        .ArmourEqpObjIndex = .Object(.ArmourEqpSlot).ObjIndex
        UserList(userIndex).Char.Body = ObjData(.ArmourEqpObjIndex).Ropaje
    Else
        Call DarCuerpoDesnudo(userIndex)
    End If
    
    If .WeaponEqpSlot > 0 Then
        .Object(.WeaponEqpSlot).Equipped = 1
        .WeaponEqpObjIndex = .Object(.WeaponEqpSlot).ObjIndex
        If .Object(.WeaponEqpSlot).ObjIndex > 0 Then UserList(userIndex).Char.WeaponAnim = ObjData(.WeaponEqpObjIndex).WeaponAnim
    Else
        UserList(userIndex).Char.WeaponAnim = NingunArma
    End If
    
    If .EscudoEqpSlot > 0 Then
        .Object(.EscudoEqpSlot).Equipped = 1
        .EscudoEqpObjIndex = .Object(.EscudoEqpSlot).ObjIndex
        UserList(userIndex).Char.ShieldAnim = ObjData(.EscudoEqpObjIndex).ShieldAnim
    Else
        UserList(userIndex).Char.ShieldAnim = NingunEscudo
    End If

    If .MunicionEqpSlot Then
        .Object(.MunicionEqpSlot).Equipped = 1
        .MunicionEqpObjIndex = .Object(.MunicionEqpSlot).ObjIndex
    End If
    
    

End With

End Sub
Public Function Insert_New_Table(ByRef name As String) As Integer
'************************************************************************
'Autor: Jose Ignacio Castelli ( Fedudok )
'Fecha: 21/7/2011
'************************************************************************
On Error GoTo Erro
    Dim iPJ As Integer
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Con.Execute "INSERT INTO `charflags` (NOMBRE) VALUES ('" & name & "')"
    Set RS = Con.Execute("SELECT * FROM `charflags` WHERE Nombre='" & name & "'")
        iPJ = RS!IndexPJ
    Set RS = Nothing

    Con.Execute "INSERT INTO `charatrib` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charbanco` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charfaccion` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charguild` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "UPDATE `charguild` SET GuildIndex=0 WHERE IndexPJ=" & iPJ
    
    Con.Execute "INSERT INTO `charhechizos` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charinit` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charinvent` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charskills` (IndexPJ) VALUES (" & iPJ & ")"
    Con.Execute "INSERT INTO `charstats` (IndexPJ) VALUES (" & iPJ & ")"
    
    Insert_New_Table = iPJ
    Exit Function
Erro:
    LogError "Insert_New_Table " & name & " " & Err.Number & " " & Err.description
End Function

#End If
