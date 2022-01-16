Attribute VB_Name = "modFileIO"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Function FileExist(ByVal File As String, _
                   Optional FileType As VbFileAttribute = vbNormal) As Boolean
    '*****************************************************************
    'Se fija si existe el archivo
    '*****************************************************************

    FileExist = LenB(dir$(File, FileType)) <> 0

End Function

Function FileRequired(ByVal File As String)

    If Not FileExist(File, vbArchive) Then
        MsgBox "Se requiere el archivo de configuraci�n " & File, vbCritical + vbOKOnly
        End
    End If

End Function

Function ValidDirectory(ByVal Path As String) As String

    If Right(Path, 1) = "\" Then
        ValidDirectory = Path
    Else
        ValidDirectory = Path & "\"
    End If
    
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a string
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 16/03/2012 - ^[GS]^
'Gets a field from a delimited string
'*****************************************************************

    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If lastPos = 0 And Pos <> 1 Then ' GSZAO, fix
        ReadField = vbNullString
        Exit Function
    End If
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
    
End Function

Public Sub CargarSpawnList()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim N As Integer, LoopC As Integer
    N = val(GetVar(pathDats & "Invokar.dat", "INIT", "NumNPCs"))
    ReDim modDeclaraciones.SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        modDeclaraciones.SpawnList(LoopC).NpcIndex = val(GetVar(pathDats & "Invokar.dat", "LIST", "NI" & LoopC))
        modDeclaraciones.SpawnList(LoopC).NpcName = GetVar(pathDats & "Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdmin(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************

    EsAdmin = (val(Administradores.GetValue("Admin", Replace(Name, "+", " "))) = 1)

End Function

Function EsDios(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************

    EsDios = (val(Administradores.GetValue("Dios", Replace(Name, "+", " "))) = 1)

End Function

Function EsSemiDios(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************

    EsSemiDios = (val(Administradores.GetValue("SemiDios", Replace(Name, "+", " "))) = 1)

End Function

Function EsGmEspecial(ByRef Name As String) As Boolean ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************

    EsGmEspecial = (val(Administradores.GetValue("Especial", Replace(Name, "+", " "))) = 1)
    
End Function

Function EsConsejero(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************

    EsConsejero = (val(Administradores.GetValue("Consejero", Replace(Name, "+", " "))) = 1)

End Function

Function EsRolesMaster(ByRef Name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************

    EsRolesMaster = (val(Administradores.GetValue("RM", Replace(Name, "+", " "))) = 1)

End Function

Public Function EsGmChar(ByRef Name As String) As Boolean ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 10/07/2012 - ^[GS]^
'Returns true if char is administrative user.
'***************************************************
    
    Dim EsGm As Boolean
    
    ' Admin?
    EsGm = EsAdmin(Name)
    ' Dios?
    If Not EsGm Then EsGm = EsDios(Name)
    ' Semidios?
    If Not EsGm Then EsGm = EsSemiDios(Name)
    ' Consejero?
    If Not EsGm Then EsGm = EsConsejero(Name)

    EsGmChar = EsGm

End Function

Public Sub LoadAdministrativeUsers() ' 0.13.3
'***************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************

    'Admines     => Admin
    'Dioses      => Dios
    'SemiDioses  => SemiDios
    'Especiales  => Especial
    'Consejeros  => Consejero
    'RoleMasters => RM

    'Si esta mierda tuviese array asociativos el c�digo ser�a tan lindo.
    Dim buf As Integer
    Dim i As Long
    Dim Name As String
       
    ' Public container
    Set Administradores = New clsIniManager
    
    ' Server ini info file
    Dim ServerIni As clsIniManager
    Set ServerIni = New clsIniManager
    
    Call ServerIni.Initialize(pathServer & fileServerIni)
       
    ' Admins
    buf = val(ServerIni.GetValue("CARGOS", "Admins"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("ADMINS", "Admin" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Admin", Name, "1")

    Next i
    
    ' Dioses
    buf = val(ServerIni.GetValue("CARGOS", "Dioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("DIOSES", "Dios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Dios", Name, "1")
        
    Next i
    
    ' Especiales
    buf = val(ServerIni.GetValue("CARGOS", "Especiales"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("ESPECIALES", "Especial" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Especial", Name, "1")
        
    Next i
    
    ' SemiDioses
    buf = val(ServerIni.GetValue("CARGOS", "SemiDioses"))
    
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("SEMIDIOSES", "SemiDios" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("SemiDios", Name, "1")
        
    Next i
    
    ' Consejeros
    buf = val(ServerIni.GetValue("CARGOS", "Consejeros"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("CONSEJEROS", "Consejero" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("Consejero", Name, "1")
        
    Next i
    
    ' RolesMasters
    buf = val(ServerIni.GetValue("CARGOS", "RolesMasters"))
        
    For i = 1 To buf
        Name = UCase$(ServerIni.GetValue("ROLESMASTERS", "RM" & i))
        
        If Left$(Name, 1) = "*" Or Left$(Name, 1) = "+" Then Name = Right$(Name, Len(Name) - 1)
        
        ' Add key
        Call Administradores.ChangeValue("RM", Name, "1")
    Next i
    
    Set ServerIni = Nothing
    
End Sub

Public Function GetCharPrivs(ByRef UserName As String) As PlayerType ' 0.13.3
'****************************************************
'Author: ZaMa
'Last Modification: 10/07/2012 - ^[GS]^
'Reads the user's charfile and retrieves its privs.
'***************************************************

    Dim Privs As PlayerType

    If EsAdmin(UserName) Then
        Privs = PlayerType.Admin
        
    ElseIf EsDios(UserName) Then
        Privs = PlayerType.Dios

    ElseIf EsSemiDios(UserName) Then
        Privs = PlayerType.SemiDios
        
    ElseIf EsConsejero(UserName) Then
        Privs = PlayerType.Consejero
    
    Else
        Privs = PlayerType.User
    End If

    GetCharPrivs = Privs

End Function

Public Function TxtDimension(ByVal Name As String) As Long
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim N As Integer, cad As String, Tam As Long
    N = FreeFile(1)
    Open Name For Input As #N
    Tam = 0
    Do While Not EOF(N)
        Tam = Tam + 1
        Line Input #N, cad
    Loop
    Close N
    TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    ReDim ForbidenNames(1 To TxtDimension(pathDats & "NombresInvalidos.txt"))
    Dim N As Integer, i As Integer
    N = FreeFile(1)
    Open pathDats & "NombresInvalidos.txt" For Input As #N
    
    For i = 1 To UBound(ForbidenNames)
        Line Input #N, ForbidenNames(i)
    Next i
    
    Close N

End Sub

Public Sub CargarHechizos()
'***************************************************
'Author: Unknownn
'Last Modification: 29/04/2013 - ^[GS]^
'
'***************************************************

On Error GoTo ErrHandler


    Dim Hechizo As Integer
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(pathDats & "Hechizos.dat")
    
    'obtiene el numero de hechizos
    NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))
    
    ReDim Hechizos(1 To NumeroHechizos) As tHechizo
    
    frmCargando.pCargar.min = 0
    frmCargando.pCargar.max = NumeroHechizos
    frmCargando.pCargar.Value = 0
    
    'Llena la lista
    For Hechizo = 1 To NumeroHechizos
        With Hechizos(Hechizo)
            .Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
            
            .GrhIndex = val(Leer.GetValue("Hechizo" & Hechizo, "GrhIndex")) ' GSZAO
            If .GrhIndex = 0 Then .GrhIndex = 609 ' Imagen de Hechizo "generico"
            
            .desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
            .PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
            
            .HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
            .targetMSG = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
            .PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
            
            .ExclusivoClase = val(Leer.GetValue("Hechizo" & Hechizo, "ExclusivoClase")) ' GSZAO
            .ExclusivoRaza = val(Leer.GetValue("Hechizo" & Hechizo, "ExclusivoRaza")) ' GSZAO
            
            .tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
            .WAV = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
            .FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
            
            .PartIndex = val(Leer.GetValue("Hechizo" & Hechizo, "Particle"))
            
            .loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
            
        '    .Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
            
            .SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
            .MinHp = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
            .MaxHp = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
            
            .SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
            .MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
            .MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
            
            .SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
            .MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
            .MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
            
            .SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
            .MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
            .MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
            
            .SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
            .MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
            .MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
            
            .SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
            .MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
            .MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
            
            .SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
            .MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
            .MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
            
            .SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
            .MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
            .MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
            
            
            .Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
            .Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
            .Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
            .RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
            .RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
            .RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
            
            
            .CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
            .Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
            .Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
            .RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
            .Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
            .Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
            
            .Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
            .Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
            
            .Warp = val(Leer.GetValue("Hechizo" & Hechizo, "Warp"))
            
            .Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
            .NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
            .cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
            .Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
            
        '    .Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
        '    .ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
            
            .MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
            .ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
            
            ' 30/9/03
            .StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
            
            .Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
            frmCargando.pCargar.Value = frmCargando.pCargar.Value + 1
            
            .NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
            .StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
            
            ' GSZAO
            .ReqObjNum = val(Leer.GetValue("Hechizo" & Hechizo, "ReqObjNum"))
            If .ReqObjNum > 0 Then
                Dim sTemp As String
                Dim LoopC As Byte
                If .ReqObjNum > MAX_INVENTORY_SLOTS Then .ReqObjNum = MAX_INVENTORY_SLOTS
                ReDim .ReqObj(.ReqObjNum) As UserOBJ
                For LoopC = 1 To .ReqObjNum
                    sTemp = Leer.GetValue("Hechizo" & Hechizo, "ReqObj" & LoopC)
                    .ReqObj(LoopC).ObjIndex = val(ReadField(1, sTemp, 45))
                    .ReqObj(LoopC).Equipped = val(ReadField(2, sTemp, 45))
                    .ReqObj(LoopC).Amount = val(ReadField(3, sTemp, 45))
                Next
            End If
            ' GSZAO
        End With
    Next Hechizo
    
    Set Leer = Nothing
    
    Exit Sub

ErrHandler:
    MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.description
 
End Sub

Sub LoadMotd()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim i As Integer
    
    MaxLines = val(GetVar(pathDats & "Motd.ini", "INIT", "NumLines"))
    
    ReDim MOTD(1 To MaxLines)
    For i = 1 To MaxLines
        MOTD(i).texto = GetVar(pathDats & "Motd.ini", "Motd", "Line" & i)
        MOTD(i).Formato = vbNullString
    Next i

End Sub

Public Sub DoBackUp()
'***************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************

    haciendoBK = True
   
    ' Lo saco porque elimina elementales y mascotas - Maraxus
    ''''''''''''''lo pongo aca x sugernecia del yind
    'For i = 1 To LastNPC
    '    If Npclist(i).flags.NPCActive Then
    '        If Npclist(i).Contadores.TiempoExistencia > 0 Then
    '            Call MuereNpc(i, 0)
    '        End If
    '    End If
    'Next i
    '''''''''''/'lo pongo aca x sugernecia del yind
   
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
    
    Call LimpiarMundo
    Call WorldSave
    Call modGuilds.v_RutinaElecciones
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

    haciendoBK = False
    
    'Log
    On Error GoTo 0 ' GSZTEST Resume Next
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time
    Close #nfile
End Sub

Public Sub GrabarMapa(ByVal Map As Long, ByRef MAPFILE As String)
'***************************************************
'Author: Unknownn
'Last Modification: 05/09/2012 - ^[GS]^
'***************************************************

On Error GoTo 0 ' GSZTEST Resume Next
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Byte
    Dim LoopC As Long
    
    ' 0.13.3
    Dim NpcInvalido As Boolean
    Dim MapWriter As clsByteBuffer
    Dim InfWriter As clsByteBuffer
    Dim IniManager As clsIniManager
    Set MapWriter = New clsByteBuffer
    Set InfWriter = New clsByteBuffer
    Set IniManager = New clsIniManager
    
    If FileExist(MAPFILE & ".map", vbNormal) Then
        Kill MAPFILE & ".map"
    End If
    
    If FileExist(MAPFILE & ".inf", vbNormal) Then
        Kill MAPFILE & ".inf"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".Map" For Binary As FreeFileMap
    
    Call MapWriter.initializeWriter(FreeFileMap)
    
    'Open .inf file
    FreeFileInf = FreeFile
    Open MAPFILE & ".Inf" For Binary As FreeFileInf
    
    Call InfWriter.initializeWriter(FreeFileInf)
    
    'map Header
    Call MapWriter.putInteger(MapInfo(Map).MapVersion)
        
    Call MapWriter.putString(MiCabecera.desc, False)
    Call MapWriter.putLong(MiCabecera.crc)
    Call MapWriter.putLong(MiCabecera.MagicWord)
    
    Call MapWriter.putDouble(0)
    
    'inf Header
    Call InfWriter.putDouble(0)
    Call InfWriter.putInteger(0)
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(Map, X, Y)
                ByFlags = 0
                
                If .Blocked Then ByFlags = ByFlags Or 1
                If .Graphic(2) Then ByFlags = ByFlags Or 2
                If .Graphic(3) Then ByFlags = ByFlags Or 4
                If .Graphic(4) Then ByFlags = ByFlags Or 8
                If .trigger Then ByFlags = ByFlags Or 16
                
                Call MapWriter.putByte(ByFlags)
                
                Call MapWriter.putInteger(.Graphic(1))
                
                For LoopC = 2 To 4
                    If .Graphic(LoopC) Then Call MapWriter.putInteger(.Graphic(LoopC))
                Next LoopC
                
                If .trigger Then Call MapWriter.putInteger(CInt(.trigger))
                
                '.inf file
                ByFlags = 0
                
                If .ObjInfo.ObjIndex > 0 Then
                   If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otFogata Then
                        .ObjInfo.ObjIndex = 0
                        .ObjInfo.Amount = 0
                    End If
                End If
    
                If .TileExit.Map Then ByFlags = ByFlags Or 1
                
                ' No hacer backup de los NPCs inv�lidos (Pretorianos, Mascotas, Invocados y Centinela)
                If .NpcIndex Then
                    NpcInvalido = (Npclist(.NpcIndex).NPCtype = eNPCType.Pretoriano) Or (Npclist(.NpcIndex).MaestroUser > 0) Or EsCentinela(.NpcIndex)
                    
                    If Not NpcInvalido Then ByFlags = ByFlags Or 2
                End If
                
                If .ObjInfo.ObjIndex Then ByFlags = ByFlags Or 4
                
                Call InfWriter.putByte(ByFlags)
                
                If .TileExit.Map Then
                    Call InfWriter.putInteger(.TileExit.Map)
                    Call InfWriter.putInteger(.TileExit.X)
                    Call InfWriter.putInteger(.TileExit.Y)
                End If
                
                ' 0.13.3
                If .NpcIndex And Not NpcInvalido Then _
                    Call InfWriter.putInteger(Npclist(.NpcIndex).Numero)
                
                If .ObjInfo.ObjIndex Then
                    Call InfWriter.putInteger(.ObjInfo.ObjIndex)
                    Call InfWriter.putInteger(.ObjInfo.Amount)
                End If
                
                NpcInvalido = False
            End With
        Next X
    Next Y
    
    Call MapWriter.saveBuffer
    Call InfWriter.saveBuffer
    
    'Close .map file
    Close FreeFileMap

    'Close .inf file
    Close FreeFileInf
    
    Set MapWriter = Nothing
    Set InfWriter = Nothing

    With MapInfo(Map)
        'write .dat file
        Call IniManager.ChangeValue("Mapa" & Map, "Name", .Name)
        Call IniManager.ChangeValue("Mapa" & Map, "MusicNum", .Music)
        Call IniManager.ChangeValue("Mapa" & Map, "MagiaSinefecto", .MagiaSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InviSinEfecto", .InviSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "ResuSinEfecto", .ResuSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "StartPos", .StartPos.Map & "-" & .StartPos.X & "-" & .StartPos.Y)
        Call IniManager.ChangeValue("Mapa" & Map, "OnDeathGoTo", .OnDeathGoTo.Map & "-" & .OnDeathGoTo.X & "-" & .OnDeathGoTo.Y) ' 0.13.3
    
        Call IniManager.ChangeValue("Mapa" & Map, "Terreno", TerrainByteToString(.Terreno))
        Call IniManager.ChangeValue("Mapa" & Map, "Zona", .Zona)
        Call IniManager.ChangeValue("Mapa" & Map, "Restringir", RestrictByteToString(.Restringir))
        Call IniManager.ChangeValue("Mapa" & Map, "BackUp", str$(.Backup))
    
        If .Pk Then
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "0")
        Else
            Call IniManager.ChangeValue("Mapa" & Map, "Pk", "1")
        End If
        
        Call IniManager.ChangeValue("Mapa" & Map, "OcultarSinEfecto", .OcultarSinEfecto)
        Call IniManager.ChangeValue("Mapa" & Map, "InvocarSinEfecto", .InvocarSinEfecto)
        ' 0.13.3
        Call IniManager.ChangeValue("Mapa" & Map, "NoEncriptarMP", .NoEncriptarMP)
        Call IniManager.ChangeValue("Mapa" & Map, "RoboNpcsPermitido", .RoboNpcsPermitido)
    
        Call IniManager.DumpFile(MAPFILE & ".dat")

    End With
    
    Set IniManager = Nothing
End Sub



Sub LoadHerreriaArmas()
'***************************************************
'Author: Unknownn
'Last Modification: 10/08/2014 - ^[GS]^
'
'***************************************************

    Dim N As Integer
    Dim lc As Integer
    Dim sArch As String
    
    sArch = pathDats & "ConstHerreroArmas.dat"
    
    N = val(GetVar(sArch, "INIT", "NumArmas"))
    ReDim Preserve lHerreroArmas(1 To N) As Integer
    For lc = 1 To N
        lHerreroArmas(lc) = val(GetVar(sArch, "Arma" & lc, "Index"))
    Next lc

End Sub

Sub LoadHerreriaArmaduras()
'***************************************************
'Author: Unknownn
'Last Modification: 10/08/2014 - ^[GS]^
'
'***************************************************

    Dim N As Integer
    Dim lc As Integer
    Dim sArch As String
    
    sArch = pathDats & "ConstHerreroArmaduras.dat"
    
    N = val(GetVar(sArch, "INIT", "NumArmaduras"))
    ReDim Preserve lHerreroArmaduras(1 To N) As Integer
    For lc = 1 To N
        lHerreroArmaduras(lc) = val(GetVar(sArch, "Armadura" & lc, "Index"))
    Next lc

End Sub

Sub LoadBalance()
'***************************************************
'Author: Unknownn
'Last Modification: 15/04/2010
'15/04/2010: ZaMa - Agrego recompensas faccionarias.
'***************************************************

    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        With ModClase(i)
            .Evasion = val(GetVar(pathDats & "Balance.dat", "MODEVASION", ListaClases(i)))
            .AtaqueArmas = val(GetVar(pathDats & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
            .AtaqueProyectiles = val(GetVar(pathDats & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
            .AtaqueWrestling = val(GetVar(pathDats & "Balance.dat", "MODATAQUEWRESTLING", ListaClases(i)))
            .Da�oArmas = val(GetVar(pathDats & "Balance.dat", "MODDA�OARMAS", ListaClases(i)))
            .Da�oProyectiles = val(GetVar(pathDats & "Balance.dat", "MODDA�OPROYECTILES", ListaClases(i)))
            .Da�oWrestling = val(GetVar(pathDats & "Balance.dat", "MODDA�OWRESTLING", ListaClases(i)))
            .Escudo = val(GetVar(pathDats & "Balance.dat", "MODESCUDO", ListaClases(i)))
        End With
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        With ModRaza(i)
            .Fuerza = val(GetVar(pathDats & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
            .Agilidad = val(GetVar(pathDats & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
            .Inteligencia = val(GetVar(pathDats & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
            .Carisma = val(GetVar(pathDats & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
            .Constitucion = val(GetVar(pathDats & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
        End With
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(pathDats & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribuci�n de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(pathDats & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i
    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(pathDats & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(pathDats & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = val(GetVar(pathDats & "Balance.dat", "PARTY", "ExponenteNivelParty"))
    
    ' Recompensas faccionarias
    For i = 1 To NUM_RANGOS_FACCION
        RecompensaFacciones(i - 1) = val(GetVar(pathDats & "Balance.dat", "RECOMPENSAFACCION", "Rango" & i))
    Next i
    
End Sub

Sub LoadCarpinteria()
'***************************************************
'Author: Unknownn
'Last Modification: 10/08/2014 - ^[GS]^
'
'***************************************************

    Dim N As Integer
    Dim lc As Integer
    Dim sArch As String
    
    sArch = pathDats & "ConstCarpintero.dat"
    
    N = val(GetVar(sArch, "INIT", "NumObjs"))
    ReDim Preserve lCarpintero(1 To N) As Integer
    For lc = 1 To N
        lCarpintero(lc) = val(GetVar(sArch, "Obj" & lc, "Index"))
    Next lc

End Sub



Sub LoadOBJData()
'***************************************************
'Author: Unknownn
'Last Modification: 09/06/2013 - ^[GS]^
'***************************************************

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'���� NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendr� que ver
'con migo. Para leer desde el OBJ.DAT se deber� usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

On Error GoTo ErrHandler

    '*****************************************************************
    'Carga la lista de objetos
    '*****************************************************************
    Dim Object As Integer
    Dim Leer As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(pathDats & "Obj.dat")
    
    'obtiene el numero de obj
    NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))
    
    frmCargando.pCargar.min = 0
    frmCargando.pCargar.max = NumObjDatas
    frmCargando.pCargar.Value = 0
    
    
    ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
    ReDim Preserve ObjListNames(1 To NumObjDatas) As String
    
    ' GSZAO
    ObjMatrimonio1 = 0
    ObjMatrimonio2 = 0
    ' GSZAO
    
    'Llena la lista
    For Object = 1 To NumObjDatas
        With ObjData(Object)
            .Name = Leer.GetValue("OBJ" & Object, "Name")
            
            If LenB(.Name) <> 0 Then
                ObjListNames(Object) = QuitarTildes(.Name) ' GSZAO
            Else
                ObjListNames(Object) = vbNullString
            End If
            
            'Pablo (ToxicWaste) Log de Objetos.
            .Log = val(Leer.GetValue("OBJ" & Object, "Log"))
            .NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
            '07/09/07
            
            .GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
            If .GrhIndex = 0 Then
                .GrhIndex = .GrhIndex
            End If
            
            .OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
            .Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
            
            .Respawn = val(Leer.GetValue("OBJ" & Object, "Respawn")) ' GSZAO
            .Bloqueado = val(Leer.GetValue("OBJ" & Object, "Bloqueado")) ' GSZAO
            
            Select Case .OBJType
                Case eOBJType.otArmadura
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                
                Case eOBJType.otESCUDO
                    .ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otCASCO
                    .CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otWeapon
                    .WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
                    .Apu�ala = val(Leer.GetValue("OBJ" & Object, "Apu�ala"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
                    .Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
                    .StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
                    .StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
                    .Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
                    
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    'Helios Fundir
                    .Fundir = val(Leer.GetValue("OBJ" & Object, "Fundir"))
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                    
                    .WeaponRazaEnanaAnim = val(Leer.GetValue("OBJ" & Object, "RazaEnanaAnim"))
                
                Case eOBJType.otInstrumentos
                    .Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
                    .Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
                    .Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
                    'Pablo (ToxicWaste)
                    .Real = val(Leer.GetValue("OBJ" & Object, "Real"))
                    .Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
                
                Case eOBJType.otMinerales
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    
                Case eOBJType.otPasaje  ' GSZAO
                    .Pasaje.Map = CInt(ReadField(1, Leer.GetValue("OBJ" & Object, "Pasaje"), 45))
                    .Pasaje.X = CInt(ReadField(2, Leer.GetValue("OBJ" & Object, "Pasaje"), 45))
                    .Pasaje.Y = CInt(ReadField(3, Leer.GetValue("OBJ" & Object, "Pasaje"), 45))
                
                Case eOBJType.otDestruible ' GSZAO
                    ' no requiere nada en especial
                
                Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
                    .IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
                    .IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
                    .IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
                
                Case otPociones
                    .TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
                    .MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
                    .MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
                    .DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
                
                Case eOBJType.otBarcos
                    .MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                
                Case eOBJType.otFlechas
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    .Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
                    .Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
                    
                Case eOBJType.otAnillo 'Pablo (ToxicWaste)
                    .LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
                    .LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
                    .LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
                    .SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
                    .MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
                    .MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
                    
                Case eOBJType.otTeleport
                    .Radio = val(Leer.GetValue("OBJ" & Object, "Radio"))
                    
                Case eOBJType.otMochilas
                    .MochilaType = val(Leer.GetValue("OBJ" & Object, "MochilaType"))
                    
                Case eOBJType.otForos
                    Call AddForum(Leer.GetValue("OBJ" & Object, "ID"))
                    
                Case eOBJType.otMatrimonio
                    ' GSZAO
                    If val(Leer.GetValue("OBJ" & Object, "Matrimonio")) = 1 Then
                        ObjMatrimonio1 = Object ' para divorciarse
                    ElseIf val(Leer.GetValue("OBJ" & Object, "Matrimonio")) = 2 Then
                        ObjMatrimonio2 = Object ' para casarse
                    End If
                    ' GSZAO
                    
            End Select
            
            .Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
            .HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
            
            .LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
            
            .MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
            
            .MaxHp = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
            .MinHp = val(Leer.GetValue("OBJ" & Object, "MinHP"))
            
            .Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
            .Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
            
            .MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
            .MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
            
            .MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
            .MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
            .def = (.MinDef + .MaxDef) / 2
            
            .RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
            .RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
            .RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
            .RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
            .RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
            
            .Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
            
            .NoLimpiar = val(Leer.GetValue("OBJ" & Object, "NoLimpiar")) ' GSZAO
            
            .Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
            If .Cerrada = 1 Then
                .Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
                .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            End If
            
            'Puertas y llaves
            .clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
            
            .texto = Leer.GetValue("OBJ" & Object, "Texto")
            .GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
            
            .NoAgarrable = val(Leer.GetValue("OBJ" & Object, "NoAgarrable")) ' GSZAO
            .ForoID = Leer.GetValue("OBJ" & Object, "ID")
            
            .Acuchilla = val(Leer.GetValue("OBJ" & Object, "Acuchilla"))
            
            .Guante = val(Leer.GetValue("OBJ" & Object, "Guante"))
            
            'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
            Dim i As Integer
            Dim N As Integer
            Dim S As String
            For i = 1 To NUMCLASES
                S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
                N = 1
                Do While LenB(S) > 0 And UCase$(ListaClases(N)) <> S
                    N = N + 1
                Loop
                .ClaseProhibida(i) = IIf(LenB(S) > 0, N, 0)
            Next i
            
            .DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
            .DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
            
            .SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
            
            If .SkCarpinteria > 0 Then .Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
                .MaderaElfica = val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
            
            'Bebidas
            .MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
            
            .NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
            
            ' 0.13.5
            .NoSeTira = val(Leer.GetValue("OBJ" & Object, "NoSeTira"))
            .NoRobable = val(Leer.GetValue("OBJ" & Object, "NoRobable"))
            .NoComerciable = val(Leer.GetValue("OBJ" & Object, "NoComerciable"))
            .Intransferible = val(Leer.GetValue("OBJ" & Object, "Intransferible"))
            .ImpideParalizar = CByte(val(Leer.GetValue("OBJ" & Object, "ImpideParalizar")))
            .ImpideInmobilizar = CByte(val(Leer.GetValue("OBJ" & Object, "ImpideInmobilizar")))
            .ImpideAturdir = CByte(val(Leer.GetValue("OBJ" & Object, "ImpideAturdir")))
            .ImpideCegar = CByte(val(Leer.GetValue("OBJ" & Object, "ImpideCegar")))
            
            .Upgrade = val(Leer.GetValue("OBJ" & Object, "Upgrade"))
            
            frmCargando.pCargar.Value = frmCargando.pCargar.Value + 1
        End With
    Next Object
    
    
    Set Leer = Nothing
    
    ' Inicializo los foros faccionarios
    Call AddForum(FORO_CAOS_ID)
    Call AddForum(FORO_REAL_ID)
    
    Exit Sub

ErrHandler:
    MsgBox "Error cargando objetos N�" & Err.Number & ": " & Err.description & vbCrLf & "Objeto: " & Object, vbCritical + vbOKOnly
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source & " - Cargando el OBJ: " & Object)

End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
'*************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'11/19/2009: Pato - Load the EluSkills and ExpSkills
'*************************************************
Dim LoopC As Long

With UserList(UserIndex)
    With .Stats
        For LoopC = 1 To NUMATRIBUTOS
            .UserAtributos(LoopC) = CInt(UserFile.GetValue("ATRIBUTOS", "AT" & LoopC))
            .UserAtributosBackUP(LoopC) = .UserAtributos(LoopC)
        Next LoopC
        
        For LoopC = 1 To NUMSKILLS
            .UserSkills(LoopC) = val(UserFile.GetValue("SKILLS", "SK" & LoopC))
            .EluSkills(LoopC) = val(UserFile.GetValue("SKILLS", "ELUSK" & LoopC))
            .ExpSkills(LoopC) = val(UserFile.GetValue("SKILLS", "EXPSK" & LoopC))
        Next LoopC
        
        For LoopC = 1 To MAXUSERHECHIZOS
            .UserHechizos(LoopC) = val(UserFile.GetValue("Hechizos", "H" & LoopC))
        Next LoopC
        
        .GLD = CLng(UserFile.GetValue("STATS", "GLD"))
        .Banco = CLng(UserFile.GetValue("STATS", "BANCO"))
        
        .MaxHp = CInt(UserFile.GetValue("STATS", "MaxHP"))
        .MinHp = CInt(UserFile.GetValue("STATS", "MinHP"))
        
        .MinSta = CInt(UserFile.GetValue("STATS", "MinSTA"))
        .MaxSta = CInt(UserFile.GetValue("STATS", "MaxSTA"))
        
        .MaxMAN = CInt(UserFile.GetValue("STATS", "MaxMAN"))
        .MinMAN = CInt(UserFile.GetValue("STATS", "MinMAN"))
        
        .MaxHIT = CInt(UserFile.GetValue("STATS", "MaxHIT"))
        .MinHIT = CInt(UserFile.GetValue("STATS", "MinHIT"))
        
        .MaxAGU = CByte(UserFile.GetValue("STATS", "MaxAGU"))
        .MinAGU = CByte(UserFile.GetValue("STATS", "MinAGU"))
        
        .MaxHam = CByte(UserFile.GetValue("STATS", "MaxHAM"))
        .MinHam = CByte(UserFile.GetValue("STATS", "MinHAM"))
        
        .SkillPts = CInt(UserFile.GetValue("STATS", "SkillPtsLibres"))
        
        .Exp = CDbl(UserFile.GetValue("STATS", "EXP"))
        .ELU = CLng(UserFile.GetValue("STATS", "ELU"))
        .ELV = CByte(UserFile.GetValue("STATS", "ELV"))
        
        
        .UsuariosMatados = CLng(UserFile.GetValue("MUERTES", "UserMuertes"))
        .NPCsMuertos = CInt(UserFile.GetValue("MUERTES", "NpcsMuertes"))
    End With
    
    With .flags
        If CByte(UserFile.GetValue("CONSEJO", "PERTENECE")) Then .Privilegios = .Privilegios Or PlayerType.RoyalCouncil
        
        If CByte(UserFile.GetValue("CONSEJO", "PERTENECECAOS")) Then .Privilegios = .Privilegios Or PlayerType.ChaosCouncil
    End With
End With
End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
'***************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************

    With UserList(UserIndex).Reputacion
        .AsesinoRep = val(UserFile.GetValue("REP", "Asesino"))
        .BandidoRep = val(UserFile.GetValue("REP", "Bandido"))
        .BurguesRep = val(UserFile.GetValue("REP", "Burguesia"))
        .LadronesRep = val(UserFile.GetValue("REP", "Ladrones"))
        .NobleRep = val(UserFile.GetValue("REP", "Nobles"))
        .PlebeRep = val(UserFile.GetValue("REP", "Plebe"))
        .Promedio = val(UserFile.GetValue("REP", "Promedio"))
    End With
    
End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsIniManager)
'*************************************************
'Author: Unknownn
'Last Modification: 08/05/2013 - ^[GS]^
'Loads the Users records
'
'*************************************************
    Dim LoopC As Long
    Dim ln As String
    
    With UserList(UserIndex)
        With .fAccion
            .ArmadaReal = CByte(UserFile.GetValue("FACCIONES", "EjercitoReal"))
            .FuerzasCaos = CByte(UserFile.GetValue("FACCIONES", "EjercitoCaos"))
            .CiudadanosMatados = CLng(UserFile.GetValue("FACCIONES", "CiudMatados"))
            .CriminalesMatados = CLng(UserFile.GetValue("FACCIONES", "CrimMatados"))
            .RecibioArmaduraCaos = CByte(UserFile.GetValue("FACCIONES", "rArCaos"))
            .RecibioArmaduraReal = CByte(UserFile.GetValue("FACCIONES", "rArReal"))
            .RecibioExpInicialCaos = CByte(UserFile.GetValue("FACCIONES", "rExCaos"))
            .RecibioExpInicialReal = CByte(UserFile.GetValue("FACCIONES", "rExReal"))
            .RecompensasCaos = CLng(UserFile.GetValue("FACCIONES", "recCaos"))
            .RecompensasReal = CLng(UserFile.GetValue("FACCIONES", "recReal"))
            .Reenlistadas = CByte(UserFile.GetValue("FACCIONES", "Reenlistadas"))
            .NivelIngreso = CInt(UserFile.GetValue("FACCIONES", "NivelIngreso"))
            .FechaIngreso = UserFile.GetValue("FACCIONES", "FechaIngreso")
            .MatadosIngreso = CInt(UserFile.GetValue("FACCIONES", "MatadosIngreso"))
            .NextRecompensa = CInt(UserFile.GetValue("FACCIONES", "NextRecompensa"))
        End With
        
        With .flags
            .Muerto = CByte(UserFile.GetValue("FLAGS", "Muerto"))
            .Escondido = CByte(UserFile.GetValue("FLAGS", "Escondido"))
            
            .Hambre = CByte(UserFile.GetValue("FLAGS", "Hambre"))
            .Sed = CByte(UserFile.GetValue("FLAGS", "Sed"))
            .Desnudo = CByte(UserFile.GetValue("FLAGS", "Desnudo"))
            .Navegando = CByte(UserFile.GetValue("FLAGS", "Navegando"))
            .Envenenado = CByte(UserFile.GetValue("FLAGS", "Envenenado"))
            .Paralizado = CByte(UserFile.GetValue("FLAGS", "Paralizado"))
            
            'Matrix
            .lastMap = val(UserFile.GetValue("FLAGS", "LastMap"))
            
            .Matrimonio = UserFile.GetValue("FLAGS", "Matrimonio") ' GSZAO
            
            .FormYesNoA = 0
            .FormYesNoDE = 0
            .FormYesNoType = 0
            '.SerialHD = val(UserFile.GetValue("FLAGS", "SerialHD")) 'GSZAO
        End With
        
        If .flags.Paralizado = 1 Then
            .Counters.Paralisis = Intervalos(eIntervalos.iParalizado)
        End If
        
        
        .Counters.Pena = CLng(UserFile.GetValue("COUNTERS", "Pena"))
        .Counters.AsignedSkills = CByte(val(UserFile.GetValue("COUNTERS", "SkillsAsignados")))
        
        .email = UserFile.GetValue("CONTACTO", "Email")
        
        .Genero = UserFile.GetValue("INIT", "Genero")
        .clase = UserFile.GetValue("INIT", "Clase")
        .raza = UserFile.GetValue("INIT", "Raza")
        .Hogar = UserFile.GetValue("INIT", "Hogar")
        .Char.heading = CInt(UserFile.GetValue("INIT", "Heading"))
        
        
        With .OrigChar
            .Head = CInt(UserFile.GetValue("INIT", "Head"))
            .Body = CInt(UserFile.GetValue("INIT", "Body"))
            .WeaponAnim = CInt(UserFile.GetValue("INIT", "Arma"))
            .ShieldAnim = CInt(UserFile.GetValue("INIT", "Escudo"))
            .CascoAnim = CInt(UserFile.GetValue("INIT", "Casco"))
            
            .heading = eHeading.SOUTH
        End With
        
        If .flags.Muerto = 0 Then
            .Char = .OrigChar
        Else
            .Char.Body = iCuerpoMuerto
            .Char.Head = iCabezaMuerto
            .Char.WeaponAnim = NingunArma
            .Char.ShieldAnim = NingunEscudo
            .Char.CascoAnim = NingunCasco
        End If
        
        
        .desc = UserFile.GetValue("INIT", "Desc")
        
        .Pos.Map = CInt(ReadField(1, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.X = CInt(ReadField(2, UserFile.GetValue("INIT", "Position"), 45))
        .Pos.Y = CInt(ReadField(3, UserFile.GetValue("INIT", "Position"), 45))
        
        .Invent.NroItems = CInt(UserFile.GetValue("Inventory", "CantidadItems"))
        
        '[KEVIN]--------------------------------------------------------------------
        '***********************************************************************************
        .BancoInvent.NroItems = CInt(UserFile.GetValue("BancoInventory", "CantidadItems"))
        'Lista de objetos del banco
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            ln = UserFile.GetValue("BancoInventory", "Obj" & LoopC) ' GSZAO
            .BancoInvent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .BancoInvent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
            .BancoInvent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
        Next LoopC
        '------------------------------------------------------------------------------------
        '[/KEVIN]*****************************************************************************
        
        
        'Lista de objetos
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            ln = UserFile.GetValue("Inventory", "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = CInt(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = CInt(ReadField(2, ln, 45))
            .Invent.Object(LoopC).Equipped = CByte(ReadField(3, ln, 45))
        Next LoopC
        
        'Obtiene el indice-objeto del arma
        .Invent.WeaponEqpSlot = CByte(UserFile.GetValue("Inventory", "WeaponEqpSlot"))
        If .Invent.WeaponEqpSlot > 0 Then
            .Invent.WeaponEqpObjIndex = .Invent.Object(.Invent.WeaponEqpSlot).ObjIndex
        End If
        
        'Obtiene el indice-objeto del armadura
        .Invent.ArmourEqpSlot = CByte(UserFile.GetValue("Inventory", "ArmourEqpSlot"))
        If .Invent.ArmourEqpSlot > 0 Then
            .Invent.ArmourEqpObjIndex = .Invent.Object(.Invent.ArmourEqpSlot).ObjIndex
            .flags.Desnudo = 0
        Else
            .flags.Desnudo = 1
        End If
        
        'Obtiene el indice-objeto del escudo
        .Invent.EscudoEqpSlot = CByte(UserFile.GetValue("Inventory", "EscudoEqpSlot"))
        If .Invent.EscudoEqpSlot > 0 Then
            .Invent.EscudoEqpObjIndex = .Invent.Object(.Invent.EscudoEqpSlot).ObjIndex
        End If
        
        'Obtiene el indice-objeto del casco
        .Invent.CascoEqpSlot = CByte(UserFile.GetValue("Inventory", "CascoEqpSlot"))
        If .Invent.CascoEqpSlot > 0 Then
            .Invent.CascoEqpObjIndex = .Invent.Object(.Invent.CascoEqpSlot).ObjIndex
        End If
        
        'Obtiene el indice-objeto barco
        .Invent.BarcoSlot = CByte(UserFile.GetValue("Inventory", "BarcoSlot"))
        If .Invent.BarcoSlot > 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(.Invent.BarcoSlot).ObjIndex
        End If
        
        'Obtiene el indice-objeto municion
        .Invent.MunicionEqpSlot = CByte(UserFile.GetValue("Inventory", "MunicionSlot"))
        If .Invent.MunicionEqpSlot > 0 Then
            .Invent.MunicionEqpObjIndex = .Invent.Object(.Invent.MunicionEqpSlot).ObjIndex
        End If
        
        '[Alejo]
        'Obtiene el indice-objeto anilo
        .Invent.AnilloEqpSlot = CByte(UserFile.GetValue("Inventory", "AnilloSlot"))
        If .Invent.AnilloEqpSlot > 0 Then
            .Invent.AnilloEqpObjIndex = .Invent.Object(.Invent.AnilloEqpSlot).ObjIndex
        End If
        
        .Invent.MochilaEqpSlot = val(UserFile.GetValue("Inventory", "MochilaSlot"))
        If .Invent.MochilaEqpSlot > 0 Then
            .Invent.MochilaEqpObjIndex = .Invent.Object(.Invent.MochilaEqpSlot).ObjIndex
        End If
        
        .NroMascotas = CInt(UserFile.GetValue("MASCOTAS", "NroMascotas"))
        For LoopC = 1 To MAXMASCOTAS
            .MascotasType(LoopC) = val(UserFile.GetValue("MASCOTAS", "MAS" & LoopC))
        Next LoopC
        
        ln = UserFile.GetValue("Guild", "GUILDINDEX")
        If IsNumeric(ln) Then
            .GuildIndex = CInt(ln)
        Else
            .GuildIndex = 0
        End If
    End With

End Sub

Function GetVar(ByVal File As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim sSpaces As String ' This will hold the input that the program will retrieve
    Dim szReturn As String ' This will be the defaul value if the string is not found
      
    szReturn = vbNullString
      
    sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
      
      
    GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, File
      
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()
'***************************************************
'Author: Unknownn
'Last Modification: 16/01/2022 - ^[GS]^
'
'***************************************************

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando backup."
    
    Dim Map As Integer
    Dim tFileName As String
    
    On Error GoTo 0 ' GSZTEST GoTo man
        
        NumMaps = val(GetVar(pathDats & "Map.dat", "INIT", "NumMaps"))
        fileMapFlagName = GetVar(pathDats & "Map.dat", "INIT", "MapFlagName")
        If Len(fileMapFlagName) = 0 Then
            fileMapFlagName = fileMapFlagDefault 'default
        End If
        
        Call InitAreas
        
        frmCargando.pCargar.min = 0
        frmCargando.pCargar.max = NumMaps
        frmCargando.pCargar.Value = 0
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
        
        For Map = 1 To NumMaps
            If val(GetVar(pathMaps & fileMapFlagName & Map & ".Dat", fileMapFlagName & Map, "BackUp")) <> 0 Then
                tFileName = pathMapsSave & fileMapFlagName & Map
                If Not FileExist(tFileName & ".*") Then 'Miramos que exista al menos uno de los 3 archivos, sino lo cargamos de la carpeta de los mapas
                    tFileName = pathMaps & fileMapFlagName & Map
                End If
            Else
                tFileName = pathMaps & fileMapFlagName & Map
            End If
            
            Call CargarMapa(Map, tFileName)
            
            frmCargando.pCargar.Value = frmCargando.pCargar.Value + 1
            DoEvents
        Next Map
    
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)
 
End Sub

Sub LoadMapData()
'***************************************************
'Author: Unknownn
'Last Modification: 16/01/2022 - ^[GS]^
'
'***************************************************

    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando mapas..."
    
    Dim Map As Integer
    Dim tFileName As String
    
    On Error GoTo 0 ' GSZTEST GoTo man
        
        NumMaps = val(GetVar(pathDats & "Map.dat", "INIT", "NumMaps"))
        fileMapFlagName = GetVar(pathDats & "Map.dat", "INIT", "MapFlagName")
        If Len(fileMapFlagName) = 0 Then
            fileMapFlagName = fileMapFlagDefault 'default
        End If
        
        Call InitAreas
        
        frmCargando.pCargar.min = 0
        frmCargando.pCargar.max = NumMaps
        frmCargando.pCargar.Value = 0
        
        ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
        ReDim MapInfo(1 To NumMaps) As MapInfo
          
        For Map = 1 To NumMaps
            tFileName = pathMaps & fileMapFlagName & Map
            Call CargarMapa(Map, tFileName)
            
            frmCargando.pCargar.Value = frmCargando.pCargar.Value + 1
            DoEvents
        Next Map
    
    Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & Map & " contiene errores")
    Call LogError(Date & " " & Err.description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.source)

End Sub

Public Sub CargarMapa(ByVal Map As Long, ByRef MAPFl As String)
'***************************************************
'Author: Unknownn
'Last Modification: 27/12/2012 - ^[GS]^
'***************************************************

On Error GoTo errh

    If FileExist(MAPFl & ".map", vbArchive) = False Then
        ' GSZAO - El Mapa "no existe"!
        MapInfo(Map).MapVersion = -1 ' marcar como invalido
        Exit Sub
    End If

    Dim hFile As Integer
    Dim X As Long
    Dim Y As Long
    Dim ByFlags As Byte
    Dim npcfile As String
    Dim Leer As clsIniManager
    Dim MapReader As clsByteBuffer
    Dim InfReader As clsByteBuffer
    Dim Buff() As Byte
    
    Set MapReader = New clsByteBuffer
    Set InfReader = New clsByteBuffer
    Set Leer = New clsIniManager
    
    npcfile = pathDats & "NPCs.dat"
    
    hFile = FreeFile

    Open MAPFl & ".map" For Binary As #hFile
        Seek hFile, 1
        
        ReDim Buff(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff
    Close hFile
    
    Call MapReader.initializeReader(Buff)

    'inf
    Open MAPFl & ".inf" For Binary As #hFile
        Seek hFile, 1

        ReDim Buff(LOF(hFile) - 1) As Byte
    
        Get #hFile, , Buff
    Close hFile
    
    Call InfReader.initializeReader(Buff)
    
    'map Header
    MapInfo(Map).MapVersion = MapReader.getInteger
    
    MiCabecera.desc = MapReader.getString(Len(MiCabecera.desc))
    MiCabecera.crc = MapReader.getLong
    MiCabecera.MagicWord = MapReader.getLong
    
    Call MapReader.getDouble

    'inf Header
    Call InfReader.getDouble
    Call InfReader.getInteger

    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            With MapData(Map, X, Y)
                '.map file
                ByFlags = MapReader.getByte

                If ByFlags And 1 Then .Blocked = 1

                .Graphic(1) = MapReader.getInteger

                'Layer 2 used?
                If ByFlags And 2 Then .Graphic(2) = MapReader.getInteger

                'Layer 3 used?
                If ByFlags And 4 Then .Graphic(3) = MapReader.getInteger

                'Layer 4 used?
                If ByFlags And 8 Then .Graphic(4) = MapReader.getInteger

                'Trigger used?
                If ByFlags And 16 Then .trigger = MapReader.getInteger

                '.inf file
                ByFlags = InfReader.getByte

                If ByFlags And 1 Then
                    .TileExit.Map = InfReader.getInteger
                    .TileExit.X = InfReader.getInteger
                    .TileExit.Y = InfReader.getInteger
                End If

                If ByFlags And 2 Then
                    'Get and make NPC
                     .NpcIndex = InfReader.getInteger

                    If .NpcIndex > 0 Then
                        'Si el npc debe hacer respawn en la pos
                        'original la guardamos
                        If val(GetVar(npcfile, "NPC" & .NpcIndex, "PosOrig")) = 1 Then
                            .NpcIndex = OpenNPC(.NpcIndex)
                            Npclist(.NpcIndex).Orig.Map = Map
                            Npclist(.NpcIndex).Orig.X = X
                            Npclist(.NpcIndex).Orig.Y = Y
                        Else
                            .NpcIndex = OpenNPC(.NpcIndex)
                        End If

                        Npclist(.NpcIndex).Pos.Map = Map
                        Npclist(.NpcIndex).Pos.X = X
                        Npclist(.NpcIndex).Pos.Y = Y

                        Call MakeNPCChar(True, 0, .NpcIndex, Map, X, Y)
                    End If
                End If

                If ByFlags And 4 Then
                    'Get and make Object
                    .ObjInfo.ObjIndex = InfReader.getInteger
                    .ObjInfo.Amount = InfReader.getInteger
                    If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otDestruible Then ' GSZAO
                        .ObjInfo.ExtraLong = ObjData(.ObjInfo.ObjIndex).MaxHp
                    Else
                        .ObjInfo.ExtraLong = 0
                    End If
                End If
            End With
        Next X
    Next Y
    
    Call Leer.Initialize(MAPFl & ".dat")
    
    With MapInfo(Map)
        .Name = Leer.GetValue("Mapa" & Map, "Name")
        .Music = Leer.GetValue("Mapa" & Map, "MusicNum")
        .StartPos.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "StartPos"), 45))
        .StartPos.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "StartPos"), 45))
        .StartPos.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "StartPos"), 45))
        
        ' 0.13.3
        .OnDeathGoTo.Map = val(ReadField(1, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.X = val(ReadField(2, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        .OnDeathGoTo.Y = val(ReadField(3, Leer.GetValue("Mapa" & Map, "OnDeathGoTo"), Asc("-")))
        
        .MagiaSinEfecto = val(Leer.GetValue("Mapa" & Map, "MagiaSinEfecto"))
        .InviSinEfecto = val(Leer.GetValue("Mapa" & Map, "InviSinEfecto"))
        .ResuSinEfecto = val(Leer.GetValue("Mapa" & Map, "ResuSinEfecto"))
        
        ' 0.13.3
        .OcultarSinEfecto = val(Leer.GetValue("Mapa" & Map, "OcultarSinEfecto"))
        .InvocarSinEfecto = val(Leer.GetValue("Mapa" & Map, "InvocarSinEfecto"))
        
        ' .NoEncriptarMP = val(Leer.GetValue("Mapa" & Map, "NoEncriptarMP")) ' GSZAO - no se utiliza

        .RoboNpcsPermitido = val(Leer.GetValue("Mapa" & Map, "RoboNpcsPermitido"))
        
        If val(Leer.GetValue("Mapa" & Map, "Pk")) = 0 Then
            .Pk = True
        Else
            .Pk = False
        End If
        
        .Terreno = TerrainStringToByte(Leer.GetValue("Mapa" & Map, "Terreno"))
        .Zona = Leer.GetValue("Mapa" & Map, "Zona")
        .Restringir = RestrictStringToByte(Leer.GetValue("Mapa" & Map, "Restringir"))
        .Backup = val(Leer.GetValue("Mapa" & Map, "BACKUP"))
        
        ' WorldGrid
        Call UpdateGrid(Map)
    End With
    
#If Testeo = 1 Then
    If MaxGrid > 0 Then ' Utiliza Grid
        Dim iX As Integer
        Dim iY As Integer
        For iX = 1 To 100
            For iY = 1 To 100
                If iX = 10 Then
                    MapData(Map, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(Map, iX, iY).ObjInfo.Amount = 1
                End If
                If iX = 90 Then
                    MapData(Map, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(Map, iX, iY).ObjInfo.Amount = 1
                End If
                If iY = 10 Then
                    MapData(Map, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(Map, iX, iY).ObjInfo.Amount = 1
                End If
                If iY = 90 Then
                    MapData(Map, iX, iY).ObjInfo.ObjIndex = 1
                    MapData(Map, iX, iY).ObjInfo.Amount = 1
                End If
            Next
        Next
    End If
#End If
    
    
    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
    
    Erase Buff
Exit Sub

errh:
    Call LogError("Error cargando Mapa: " & Map & " - Pos: " & X & "," & Y & "." & Err.description)

    Set MapReader = Nothing
    Set InfReader = Nothing
    Set Leer = Nothing
End Sub

Sub LoadIntervalos()
'***************************************************
'Author: ^[GS]^
'Last Modification: 16/01/2022 - ^[GS]^
'***************************************************
On Error GoTo error

    Dim s_File As String
    s_File = pathServer & fileServerIni
    
    If FileExist(s_File, vbArchive) Then
        ' ####### Leemos los valores #######
        ' Generales
        Intervalos(eIntervalos.iWorldSave) = val(GetVar(s_File, "GENERALES", "WorldSave"))
        Intervalos(eIntervalos.iGuardarUsuarios) = val(GetVar(s_File, "GENERALES", "GuardarUsuarios")) ' 0.13.3
        Intervalos(eIntervalos.iMinutosMotd) = val(GetVar(s_File, "GENERALES", "MinutosMotd")) ' 0.13.5
        Intervalos(eIntervalos.iCerrarConexion) = val(GetVar(s_File, "GENERALES", "CerrarConexion"))
        Intervalos(eIntervalos.iCerrarConexionInactivo) = val(GetVar(s_File, "GENERALES", "CerrarConexionInactivo"))
        Intervalos(eIntervalos.iEfectoLluvia) = val(GetVar(s_File, "GENERALES", "EfectoLluvia"))
        Intervalos(eIntervalos.iLluvia) = val(GetVar(s_File, "GENERALES", "Lluvia"))
        Intervalos(eIntervalos.iLluviaDeORO) = val(GetVar(s_File, "GENERALES", "LluviaDeORO"))
        Intervalos(eIntervalos.iReproducirFXMapas) = val(GetVar(s_File, "GENERALES", "IntervaloWAVFX"))
        ' Estado
        Intervalos(eIntervalos.iSanarSinDescansar) = val(GetVar(s_File, "ESTADO", "SanarSinDescansar"))
        Intervalos(eIntervalos.iStaminaSinDescansar) = val(GetVar(s_File, "ESTADO", "StaminaSinDescansar"))
        Intervalos(eIntervalos.iSanarDescansando) = val(GetVar(s_File, "ESTADO", "SanarDescansando"))
        Intervalos(eIntervalos.iStaminaDescansando) = val(GetVar(s_File, "ESTADO", "StaminaDescansando"))
        Intervalos(eIntervalos.iSed) = val(GetVar(s_File, "ESTADO", "Sed"))
        Intervalos(eIntervalos.iHambre) = val(GetVar(s_File, "ESTADO", "Hambre"))
        Intervalos(eIntervalos.iVeneno) = val(GetVar(s_File, "ESTADO", "Veneno"))
        Intervalos(eIntervalos.iParalizado) = val(GetVar(s_File, "ESTADO", "Paralizado"))
        Intervalos(eIntervalos.iInvisible) = val(GetVar(s_File, "ESTADO", "Invisible"))
        Intervalos(eIntervalos.iOculto) = val(GetVar(s_File, "ESTADO", "Oculto"))
        Intervalos(eIntervalos.iFrio) = val(GetVar(s_File, "ESTADO", "Frio"))
        Intervalos(eIntervalos.iInvocacion) = val(GetVar(s_File, "ESTADO", "Invocacion"))
        ' NPC's
        Intervalos(eIntervalos.iNPCPuedeAtacar) = val(GetVar(s_File, "NPC", "PuedeAtacar"))
        Intervalos(eIntervalos.iNPCPuedeUsarAI) = val(GetVar(s_File, "NPC", "PuedeUsarAI"))
        ' Cliente
        Intervalos(eIntervalos.iPuedeAtacar) = val(GetVar(s_File, "CLIENTE", "PuedeAtacar"))
        Intervalos(eIntervalos.iPuedeAtacarConFlechas) = val(GetVar(s_File, "CLIENTE", "PuedeAtacarConFlechas"))
        Intervalos(eIntervalos.iPuedeAtacarConHechizos) = val(GetVar(s_File, "CLIENTE", "PuedeAtacarConHechizos"))
        Intervalos(eIntervalos.iPuedeUsarItem) = val(GetVar(s_File, "CLIENTE", "PuedeUsarItem"))
        Intervalos(eIntervalos.iPuedeUsarPocion) = val(GetVar(s_File, "CLIENTE", "PuedeUsarPocion"))
        Intervalos(eIntervalos.iPuedeTrabajar) = val(GetVar(s_File, "CLIENTE", "PuedeTrabajar"))
        Intervalos(eIntervalos.iComboMagiaGolpe) = val(GetVar(s_File, "CLIENTE", "ComboMagiaGolpe"))
        Intervalos(eIntervalos.iComboGolpeMagia) = val(GetVar(s_File, "CLIENTE", "ComboGolpeMagia"))
        
    End If
    
    ' ####### Corregimos los valores por defecto en caso de un mal valor #######
    ' Generales
    If Intervalos(eIntervalos.iWorldSave) < 60 Then Intervalos(eIntervalos.iWorldSave) = 60
    If Intervalos(eIntervalos.iGuardarUsuarios) < 30 Then Intervalos(eIntervalos.iGuardarUsuarios) = 30
    If Intervalos(eIntervalos.iMinutosMotd) < 20 And Intervalos(eIntervalos.iMinutosMotd) <> 0 Then Intervalos(eIntervalos.iMinutosMotd) = 20
    If Intervalos(eIntervalos.iCerrarConexion) < 0 Then Intervalos(eIntervalos.iCerrarConexion) = 1
    If Intervalos(eIntervalos.iCerrarConexionInactivo) < 20 Then Intervalos(eIntervalos.iCerrarConexionInactivo) = 900 ' 15 minutos
    If Intervalos(eIntervalos.iEfectoLluvia) < 0 And Intervalos(eIntervalos.iEfectoLluvia) <> 0 Or Intervalos(eIntervalos.iEfectoLluvia) > 32 Then Intervalos(eIntervalos.iEfectoLluvia) = 32
    If Intervalos(eIntervalos.iLluvia) < 0 And Intervalos(eIntervalos.iLluvia) <> 0 Then Intervalos(eIntervalos.iLluvia) = 30
    If Intervalos(eIntervalos.iLluviaDeORO) < 0 And Intervalos(eIntervalos.iLluviaDeORO) <> 0 Then Intervalos(eIntervalos.iLluviaDeORO) = 10
    If Intervalos(eIntervalos.iReproducirFXMapas) < 0 And Intervalos(eIntervalos.iReproducirFXMapas) <> 0 Then Intervalos(eIntervalos.iReproducirFXMapas) = 5
    ' Estado
    If Intervalos(eIntervalos.iSanarSinDescansar) < 0 Then Intervalos(eIntervalos.iSanarSinDescansar) = 1600
    If Intervalos(eIntervalos.iStaminaSinDescansar) < 0 Then Intervalos(eIntervalos.iStaminaSinDescansar) = 10
    If Intervalos(eIntervalos.iSanarDescansando) < 0 Then Intervalos(eIntervalos.iSanarDescansando) = 100
    If Intervalos(eIntervalos.iStaminaDescansando) < 0 Then Intervalos(eIntervalos.iStaminaDescansando) = 5
    If Intervalos(eIntervalos.iSed) < 0 Then Intervalos(eIntervalos.iSed) = 6000
    If Intervalos(eIntervalos.iHambre) < 0 Then Intervalos(eIntervalos.iHambre) = 6500
    If Intervalos(eIntervalos.iVeneno) < 0 Then Intervalos(eIntervalos.iVeneno) = 500
    If Intervalos(eIntervalos.iParalizado) < 0 Then Intervalos(eIntervalos.iParalizado) = 500
    If Intervalos(eIntervalos.iInvisible) < 0 Then Intervalos(eIntervalos.iInvisible) = 500
    If Intervalos(eIntervalos.iOculto) < 0 Then Intervalos(eIntervalos.iOculto) = 500
    If Intervalos(eIntervalos.iFrio) < 0 Then Intervalos(eIntervalos.iFrio) = 15
    If Intervalos(eIntervalos.iInvocacion) < 0 Then Intervalos(eIntervalos.iInvocacion) = 1001
    ' NPC's
    If Intervalos(eIntervalos.iNPCPuedeAtacar) < 0 Then Intervalos(eIntervalos.iNPCPuedeAtacar) = 1600
    If Intervalos(eIntervalos.iNPCPuedeUsarAI) < 0 Then Intervalos(eIntervalos.iNPCPuedeUsarAI) = 400
    ' Cliente
    If Intervalos(eIntervalos.iPuedeAtacar) < 0 Then Intervalos(eIntervalos.iPuedeAtacar) = 1500
    If Intervalos(eIntervalos.iPuedeAtacarConFlechas) < 0 Then Intervalos(eIntervalos.iPuedeAtacarConFlechas) = 1400
    If Intervalos(eIntervalos.iPuedeAtacarConHechizos) < 0 Then Intervalos(eIntervalos.iPuedeAtacarConHechizos) = 1400
    If Intervalos(eIntervalos.iPuedeUsarItem) < 0 Then Intervalos(eIntervalos.iPuedeUsarItem) = 125
    If Intervalos(eIntervalos.iPuedeUsarPocion) < 0 Then Intervalos(eIntervalos.iPuedeUsarPocion) = 1200
    If Intervalos(eIntervalos.iPuedeTrabajar) < 0 Then Intervalos(eIntervalos.iPuedeTrabajar) = 700
    If Intervalos(eIntervalos.iComboMagiaGolpe) < 0 Then Intervalos(eIntervalos.iComboMagiaGolpe) = 1000
    If Intervalos(eIntervalos.iComboGolpeMagia) < 0 Then Intervalos(eIntervalos.iComboGolpeMagia) = 1000

    ' ####### Aplicamos los intervalos que lo requieran #######
    ' Generales
    frmMain.tEfectoLluvia.interval = val(Intervalos(eIntervalos.iEfectoLluvia) * 1000)
    frmMain.tFXMapas.interval = val(Intervalos(eIntervalos.iReproducirFXMapas) * 1000)
    ' NPC's
    frmMain.tNpcAtaca.interval = val(Intervalos(eIntervalos.iNPCPuedeAtacar))
    frmMain.tNpcAI.interval = val(Intervalos(eIntervalos.iNPCPuedeUsarAI))
    
    ' Constantes (GS: �Why?)
    IntervaloPuedeSerAtacado = 5000 ' Cargar desde balance.dat
    IntervaloAtacable = 60000 ' Cargar desde balance.dat
    IntervaloOwnedNpc = 18000 ' Cargar desde balance.dat
    
    Exit Sub
error:
    MsgBox "Error en LoadIntervalos: " & Err.Number & " - " & Err.description
    Call LogError("LoadIntervalos: " & Err.Number & " - " & Err.description)

End Sub

Sub LoadSini()
'***************************************************
'Author: Unknownn
'Last Modification: 10/08/2014 - ^[GS]^
'***************************************************

    Dim lTemp As Long
    Dim sTemp As String
    
    ' Init
    iniNombre = GetVar(pathServer & fileServerIni, "INIT", "Nombre") ' GSZAO
    iniWWW = GetVar(pathServer & fileServerIni, "INIT", "WWW") ' GSZAO
    iniPuerto = val(GetVar(pathServer & fileServerIni, "INIT", "Puerto"))
    iniVersion = GetVar(pathServer & fileServerIni, "INIT", "Version")
    iniOculto = val(GetVar(pathServer & fileServerIni, "INIT", "Oculto"))
    iniWorldBackup = val(GetVar(pathServer & fileServerIni, "INIT", "WorldBackup"))
    iniMapaPretoriano = val(GetVar(pathServer & fileServerIni, "INIT", "MapaPretoriano"))
    
    sTemp = GetVar(pathServer & fileServerIni, "PATHS", "PathLogs") ' GSZAO
    If LenB(sTemp) Then
        pathLogs = ValidDirectory(pathServer & sTemp)
    End If
    sTemp = GetVar(pathServer & fileServerIni, "PATHS", "PathChars") ' GSZAO
    If LenB(sTemp) Then
        pathChars = ValidDirectory(pathServer & sTemp)
    End If
    sTemp = GetVar(pathServer & fileServerIni, "PATHS", "PathDats") ' GSZAO
    If LenB(sTemp) Then
        pathDats = ValidDirectory(pathServer & sTemp)
    End If
    sTemp = GetVar(pathServer & fileServerIni, "PATHS", "PathGuilds") ' GSZAO
    If LenB(sTemp) Then
        pathGuilds = ValidDirectory(pathServer & sTemp)
    End If
    sTemp = GetVar(pathServer & fileServerIni, "PATHS", "PathMaps") ' GSZAO
    If LenB(sTemp) Then
        pathMaps = ValidDirectory(pathServer & sTemp)
    End If
    sTemp = GetVar(pathServer & fileServerIni, "PATHS", "PathMapsSave") ' GSZAO
    If LenB(sTemp) Then
        pathMapsSave = ValidDirectory(pathServer & sTemp)
    End If
    
    ' Control de directorios
    If Not FileExist(pathLogs, vbDirectory) Then
        Call MkDir(pathLogs)
    End If
    If Not FileExist(pathChars, vbDirectory) Then
        Call MkDir(pathChars)
    End If
    If Not FileExist(pathMapsSave, vbDirectory) Then
        Call MkDir(pathMapsSave)
    End If
    If Not FileExist(pathGuilds, vbDirectory) Then
        Call MkDir(pathGuilds)
    End If
    If Not FileExist(pathDats, vbDirectory) Then
        MsgBox "Se requiere la carpeta de dats: " & pathDats, vbCritical + vbOKOnly
        End
    End If
    If Not FileExist(pathMaps, vbDirectory) Then
        MsgBox "Se requiere la carpeta de mapas: " & pathMaps, vbCritical + vbOKOnly
        End
    End If
    
    iniWorldGrid = GetVar(pathServer & fileServerIni, "INIT", "WorldGrid")
    If LenB(iniWorldGrid) <> 0 Then ' GSZAO - Es de uso OPCIONAL
        If LoadWorldGrid(pathDats & iniWorldGrid & ".grid") = False Then
            Call LogError("Ha ocurrido un error durante la carga de " & pathDats & iniWorldGrid & ".grid. No se puede continuar con la ejecuci�n del servidor.")
            End
        End If
        If NumMaps > 0 Then ' Los mapas ya estan cargados...
            For lTemp = 1 To NumMaps
                Call UpdateGrid(lTemp) ' Actualizamos el Grid en los mapas
            Next
        End If
    End If
    
    iniRecord = val(GetVar(pathServer & fileServerIni, "INIT", "Record"))
    
    ' Opciones
    iniDragDrop = CByte(val(GetVar(pathServer & fileServerIni, "OPCIONES", "DragDrop")))
    iniTirarOBJZonaSegura = CByte(val(GetVar(pathServer & fileServerIni, "OPCIONES", "TirarOBJZonaSegura")))
    iniMeditarRapido = IIf(GetVar(pathServer & fileServerIni, "OPCIONES", "MeditarRapido") = 1, True, False)
    iniPrivadoPorConsola = IIf(GetVar(pathServer & fileServerIni, "OPCIONES", "PrivadoPorConsola") = 1, True, False)
    iniAutoSacerdote = IIf(GetVar(pathServer & fileServerIni, "OPCIONES", "AutoSacerdote") = 1, True, False)
    iniSacerdoteCuraVeneno = IIf(GetVar(pathServer & fileServerIni, "OPCIONES", "SacerdoteCuraVeneno") = 1, True, False)
    iniNPCNoHostilesConNombre = IIf(GetVar(pathServer & fileServerIni, "OPCIONES", "NPCNoHostilesConNombre") = 1, True, False)
    iniNPCHostilesConNombre = IIf(GetVar(pathServer & fileServerIni, "OPCIONES", "NPCHostilesConNombre") = 1, True, False)
    
    'Actualizo el frmMain. / maTih.-  |  02/03/2012
    If frmMain.Visible Then frmMain.Record = CStr(iniRecord)
    
    ' Conexiones
    lTemp = val(GetVar(pathServer & fileServerIni, "CONEXION", "MaxUsuarios"))
    If iniMaxUsuarios = 0 Then
        iniMaxUsuarios = lTemp
        ReDim UserList(1 To iniMaxUsuarios) As User
    End If
    iniMultiLogin = val(GetVar(pathServer & fileServerIni, "CONEXION", "MultiLogin"))
    iniInactivo = val(GetVar(pathServer & fileServerIni, "CONEXION", "Inactivo"))
    lastSockListen = val(GetVar(pathServer & fileServerIni, "CONEXION", "LastSockListen"))
    
    ' Balance
    iniMaxNivel = CByte(GetVar(pathServer & fileServerIni, "BALANCE", "MaxNivel")) ' max 255
    If iniMaxNivel = 0 Then iniMaxNivel = 50
    iniOro = val(GetVar(pathServer & fileServerIni, "BALANCE", "Oro"))
    iniExp = val(GetVar(pathServer & fileServerIni, "BALANCE", "Exp"))
    iniTPesca = val(GetVar(pathServer & fileServerIni, "BALANCE", "Pesca"))
    iniTMineria = val(GetVar(pathServer & fileServerIni, "BALANCE", "Mineria"))
    iniTTala = val(GetVar(pathServer & fileServerIni, "BALANCE", "Tala"))
    If iniOro <= 0 Then iniOro = 1
    If iniExp <= 0 Then iniExp = 1
    If iniTPesca <= 0 Then iniTPesca = 1
    If iniTMineria <= 0 Then iniTMineria = 1
    If iniTTala <= 0 Then iniTTala = 1
    iniBilletera = val(GetVar(pathServer & fileServerIni, "BALANCE", "Billetera")) ' con 0 se deshabilita
    iniBilleteraSegura = val(GetVar(pathServer & fileServerIni, "BALANCE", "BilleteraSegura"))
    ' HappyHour
    iniHappyHourActivado = IIf(GetVar(pathServer & fileServerIni, "HAPPYHOUR", "Activado") = 1, True, False)
    For lTemp = 1 To 7
        sTemp = GetVar(pathServer & fileServerIni, "HAPPYHOUR", "Dia" & lTemp)
        HappyHourDays(lTemp).Hour = val(ReadField(1, sTemp, 45)) ' GSZAO
        HappyHourDays(lTemp).Multi = val(ReadField(2, sTemp, 45)) ' 0.13.5
        If HappyHourDays(lTemp).Hour < 0 Or HappyHourDays(lTemp).Hour > 23 Then HappyHourDays(lTemp).Hour = 20 ' Hora de 0 a 23.
        If HappyHourDays(lTemp).Multi < 0 Then HappyHourDays(lTemp).Multi = 0
    Next
    
    ' Dados
    sTemp = GetVar(pathServer & fileServerIni, "DADOS", "Fuerza")
    Dados(0).Minimo = val(ReadField(1, sTemp, 45))
    sTemp = ReadField(2, sTemp, 45)
    If LenB(sTemp) <> 0 Then
        Dados(0).Base = val(ReadField(1, sTemp, 43))
        Dados(0).Random = val(ReadField(2, sTemp, 43))
    Else
        Dados(0).Base = 0
        Dados(0).Random = 0
    End If
    sTemp = GetVar(pathServer & fileServerIni, "DADOS", "Agilidad")
    Dados(1).Minimo = val(ReadField(1, sTemp, 45))
    sTemp = ReadField(2, sTemp, 45)
    If LenB(sTemp) <> 0 Then
        Dados(1).Base = val(ReadField(1, sTemp, 43))
        Dados(1).Random = val(ReadField(2, sTemp, 43))
    Else
        Dados(1).Base = 0
        Dados(1).Random = 0
    End If
    sTemp = GetVar(pathServer & fileServerIni, "DADOS", "Inteligencia")
    Dados(2).Minimo = val(ReadField(1, sTemp, 45))
    sTemp = ReadField(2, sTemp, 45)
    If LenB(sTemp) <> 0 Then
        Dados(2).Base = val(ReadField(1, sTemp, 43))
        Dados(2).Random = val(ReadField(2, sTemp, 43))
    Else
        Dados(2).Base = 0
        Dados(2).Random = 0
    End If
    sTemp = GetVar(pathServer & fileServerIni, "DADOS", "Carisma")
    Dados(3).Minimo = val(ReadField(1, sTemp, 45))
    sTemp = ReadField(2, sTemp, 45)
    If LenB(sTemp) <> 0 Then
        Dados(3).Base = val(ReadField(1, sTemp, 43))
        Dados(3).Random = val(ReadField(2, sTemp, 43))
    Else
        Dados(3).Base = 0
        Dados(3).Random = 0
    End If
    sTemp = GetVar(pathServer & fileServerIni, "DADOS", "Constitucion")
    Dados(4).Minimo = val(ReadField(1, sTemp, 45))
    sTemp = ReadField(2, sTemp, 45)
    If LenB(sTemp) <> 0 Then
        Dados(4).Base = val(ReadField(1, sTemp, 43))
        Dados(4).Random = val(ReadField(2, sTemp, 43))
    Else
        Dados(4).Base = 0
        Dados(4).Random = 0
    End If
    
    ' Clanes
    iniCNivel = CByte(GetVar(pathServer & fileServerIni, "CLANES", "NivelMinimo"))
    iniCLiderazgo = val(GetVar(pathServer & fileServerIni, "CLANES", "LiderazgoMinimo"))
    sTemp = GetVar(pathServer & fileServerIni, "CLANES", "RequiereOBJ")
    iniCRequiereObj = val(ReadField(1, sTemp, 45))
    iniCRequiereObjCnt = val(ReadField(2, sTemp, 45))
    If iniCNivel = 0 Then iniCNivel = 25
    If iniCLiderazgo = 0 Then iniCLiderazgo = 90
    If iniCRequiereObj = 0 Then iniCRequiereObj = 0
    If iniCRequiereObjCnt = 0 Then iniCRequiereObjCnt = 0
    
    ' Meditacion
    iniFxMedChico = CByte(GetVar(pathServer & fileServerIni, "MEDITACION", "FxChico"))
    iniFxMedMediano = CByte(GetVar(pathServer & fileServerIni, "MEDITACION", "FxMediano"))
    iniFxMedGrande = CByte(GetVar(pathServer & fileServerIni, "MEDITACION", "FxGrande"))
    iniFxMedExtraGrande = CByte(GetVar(pathServer & fileServerIni, "MEDITACION", "FxExtraGrande"))
    
    ' Cliente
    iniSiempreNombres = val(GetVar(pathServer & fileServerIni, "CLIENTE", "SiempreNombres"))
    iniDiaNoche = val(GetVar(pathServer & fileServerIni, "CLIENTE", "DiaNoche"))
    iniSistemaLuces = val(GetVar(pathServer & fileServerIni, "CLIENTE", "SistemaLuces"))
    
    ' Seguridad
    iniPuedeCrearPersonajes = val(GetVar(pathServer & fileServerIni, "SEGURIDAD", "PuedeCrearPersonajes"))
    iniSoloGMs = val(GetVar(pathServer & fileServerIni, "SEGURIDAD", "SoloGMs"))
    iniTesting = val(GetVar(pathServer & fileServerIni, "SEGURIDAD", "Testing"))
    iniLogDesarrollo = val(GetVar(pathServer & fileServerIni, "SEGURIDAD", "LogDesarrollo")) ' GSZAO
    
    ' Faccion
    ArmaduraImperial1 = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraImperial1"))
    ArmaduraImperial2 = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraImperial2"))
    ArmaduraImperial3 = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraImperial3"))
    TunicaMagoImperial = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaMagoImperial"))
    TunicaMagoImperialEnanos = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaMagoImperialEnanos"))
    ArmaduraCaos1 = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraCaos1"))
    ArmaduraCaos2 = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraCaos2"))
    ArmaduraCaos3 = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraCaos3"))
    TunicaMagoCaos = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaMagoCaos"))
    TunicaMagoCaosEnanos = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaMagoCaosEnanos"))
    
    VestimentaImperialHumano = val(GetVar(pathServer & fileServerIni, "INIT", "VestimentaImperialHumano"))
    VestimentaImperialEnano = val(GetVar(pathServer & fileServerIni, "INIT", "VestimentaImperialEnano"))
    TunicaConspicuaHumano = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaConspicuaHumano"))
    TunicaConspicuaEnano = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaConspicuaEnano"))
    ArmaduraNobilisimaHumano = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraNobilisimaHumano"))
    ArmaduraNobilisimaEnano = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraNobilisimaEnano"))
    ArmaduraGranSacerdote = val(GetVar(pathServer & fileServerIni, "INIT", "ArmaduraGranSacerdote"))
    
    VestimentaLegionHumano = val(GetVar(pathServer & fileServerIni, "INIT", "VestimentaLegionHumano"))
    VestimentaLegionEnano = val(GetVar(pathServer & fileServerIni, "INIT", "VestimentaLegionEnano"))
    TunicaLobregaHumano = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaLobregaHumano"))
    TunicaLobregaEnano = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaLobregaEnano"))
    TunicaEgregiaHumano = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaEgregiaHumano"))
    TunicaEgregiaEnano = val(GetVar(pathServer & fileServerIni, "INIT", "TunicaEgregiaEnano"))
    SacerdoteDemoniaco = val(GetVar(pathServer & fileServerIni, "INIT", "SacerdoteDemoniaco"))
    
    'Intervalos
    Call LoadIntervalos ' GSZAO

    '&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    'Se agreg� en LoadBalance y en el Balance.dat
    'PorcentajeRecuperoMana = val(GetVar(pathServer & fileServerIni, "BALANCE", "PorcentajeRecuperoMana"))
    
    ''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
    Call modStatistics.Initialize
    
    ' GSZAO - Es necesario que las ciudades se puedan configurar facilmente :)
    NUMCIUDADES = val(GetVar(pathDats & "Ciudades.dat", "INIT", "MaxCiudades"))
    If NUMCIUDADES > 0 And NUMCIUDADES <= 25 Then ' maximo 25
        ReDim Ciudades(1 To NUMCIUDADES) As WorldPos
        For lTemp = 1 To NUMCIUDADES
            Ciudades(lTemp).Map = val(GetVar(pathDats & "Ciudades.dat", "CIUDAD" & lTemp, "Mapa"))
            Ciudades(lTemp).X = val(GetVar(pathDats & "Ciudades.dat", "CIUDAD" & lTemp, "X"))
            Ciudades(lTemp).Y = val(GetVar(pathDats & "Ciudades.dat", "CIUDAD" & lTemp, "Y"))
        Next
    End If
    
    ' Prisi�n
    Prision.Map = val(GetVar(pathDats & "Ciudades.dat", "PRISION", "Mapa"))
    Prision.X = val(GetVar(pathDats & "Ciudades.dat", "PRISION", "X"))
    Prision.Y = val(GetVar(pathDats & "Ciudades.dat", "PRISION", "Y"))
    
    ' Libertad de Prisi�n
    Libertad.Map = val(GetVar(pathDats & "Ciudades.dat", "LIBERTAD", "Mapa"))
    Libertad.X = val(GetVar(pathDats & "Ciudades.dat", "LIBERTAD", "X"))
    Libertad.Y = val(GetVar(pathDats & "Ciudades.dat", "LIBERTAD", "Y"))

    Set ConsultaPopular = New clsConsultasPopulares
    Call ConsultaPopular.LoadData

    ' Admins
    Call LoadAdministrativeUsers ' 0.13.3
    
    aLluviaDeOro = False ' GSZAO

End Sub

Sub WriteVar(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'***************************************************
'Author: Unknownn
'Last Modification: -
'Escribe VAR en un archivo
'***************************************************

writeprivateprofilestring Main, Var, Value, File
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer, ByVal UserFile As String, Optional ByVal SaveTimeOnline As Boolean = True)
'*************************************************
'Author: Unknownn
'Last Modification: 14/08/2014 - ^[GS]^
'Saves the Users records
'*************************************************

On Error GoTo 0 ' GSZTEST GoTo Errhandler

Dim Manager As clsIniManager
Dim Existe As Boolean

With UserList(UserIndex)

    'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
    'clase=0 es el error, porq el enum empieza de 1!!
    If .clase = 0 Or .Stats.ELV = 0 Then
        Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & .Name)
        Exit Sub
    End If
    
    Set Manager = New clsIniManager
    
    If FileExist(UserFile) Then
        Call Manager.Initialize(UserFile)
        
        If FileExist(UserFile & ".bk") Then Call Kill(UserFile & ".bk")
        Name UserFile As UserFile & ".bk"
        
        Existe = True
    End If
    
    If .flags.Mimetizado = 1 Then
        .Char.Body = .CharMimetizado.Body
        .Char.Head = .CharMimetizado.Head
        .Char.CascoAnim = .CharMimetizado.CascoAnim
        .Char.ShieldAnim = .CharMimetizado.ShieldAnim
        .Char.WeaponAnim = .CharMimetizado.WeaponAnim
        .Counters.Mimetismo = 0
        .flags.Mimetizado = 0
        ' Se fue el efecto del mimetismo, puede ser atacado por npcs
        .flags.Ignorado = False
    End If
      
    Dim LoopC As Integer
    
    Call Manager.ChangeValue("FLAGS", "Muerto", CStr(.flags.Muerto))
    Call Manager.ChangeValue("FLAGS", "Escondido", CStr(.flags.Escondido))
    Call Manager.ChangeValue("FLAGS", "Hambre", CStr(.flags.Hambre))
    Call Manager.ChangeValue("FLAGS", "Sed", CStr(.flags.Sed))
    Call Manager.ChangeValue("FLAGS", "Desnudo", CStr(.flags.Desnudo))
    Call Manager.ChangeValue("FLAGS", "Ban", CStr(.flags.Ban))
    Call Manager.ChangeValue("FLAGS", "Navegando", CStr(.flags.Navegando))
    Call Manager.ChangeValue("FLAGS", "Envenenado", CStr(.flags.Envenenado))
    Call Manager.ChangeValue("FLAGS", "Paralizado", CStr(.flags.Paralizado))
    Call Manager.ChangeValue("FLAGS", "Matrimonio", CStr(.flags.Matrimonio))
    
    'Matrix
    Call Manager.ChangeValue("FLAGS", "LastMap", CStr(.flags.lastMap))
    Call Manager.ChangeValue("FLAGS", "SerialHD", CStr(.flags.SerialHD)) 'GSZAO
    
    Call Manager.ChangeValue("CONSEJO", "PERTENECE", IIf(.flags.Privilegios And PlayerType.RoyalCouncil, "1", "0"))
    Call Manager.ChangeValue("CONSEJO", "PERTENECECAOS", IIf(.flags.Privilegios And PlayerType.ChaosCouncil, "1", "0"))
    
    Call Manager.ChangeValue("COUNTERS", "Pena", CStr(.Counters.Pena))
    Call Manager.ChangeValue("COUNTERS", "SkillsAsignados", CStr(.Counters.AsignedSkills))
    
    Call Manager.ChangeValue("FACCIONES", "EjercitoReal", CStr(.fAccion.ArmadaReal))
    Call Manager.ChangeValue("FACCIONES", "EjercitoCaos", CStr(.fAccion.FuerzasCaos))
    Call Manager.ChangeValue("FACCIONES", "CiudMatados", CStr(.fAccion.CiudadanosMatados))
    Call Manager.ChangeValue("FACCIONES", "CrimMatados", CStr(.fAccion.CriminalesMatados))
    Call Manager.ChangeValue("FACCIONES", "rArCaos", CStr(.fAccion.RecibioArmaduraCaos))
    Call Manager.ChangeValue("FACCIONES", "rArReal", CStr(.fAccion.RecibioArmaduraReal))
    Call Manager.ChangeValue("FACCIONES", "rExCaos", CStr(.fAccion.RecibioExpInicialCaos))
    Call Manager.ChangeValue("FACCIONES", "rExReal", CStr(.fAccion.RecibioExpInicialReal))
    Call Manager.ChangeValue("FACCIONES", "recCaos", CStr(.fAccion.RecompensasCaos))
    Call Manager.ChangeValue("FACCIONES", "recReal", CStr(.fAccion.RecompensasReal))
    Call Manager.ChangeValue("FACCIONES", "Reenlistadas", CStr(.fAccion.Reenlistadas))
    Call Manager.ChangeValue("FACCIONES", "NivelIngreso", CStr(.fAccion.NivelIngreso))
    Call Manager.ChangeValue("FACCIONES", "FechaIngreso", .fAccion.FechaIngreso)
    Call Manager.ChangeValue("FACCIONES", "MatadosIngreso", CStr(.fAccion.MatadosIngreso))
    Call Manager.ChangeValue("FACCIONES", "NextRecompensa", CStr(.fAccion.NextRecompensa))
    
    '�Fueron modificados los atributos del usuario?
    If Not .flags.TomoPocion Then
        For LoopC = 1 To UBound(.Stats.UserAtributos)
            Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributos(LoopC)))
        Next LoopC
    Else
        For LoopC = 1 To UBound(.Stats.UserAtributos)
            '.Stats.UserAtributos(LoopC) = .Stats.UserAtributosBackUP(LoopC)
            Call Manager.ChangeValue("ATRIBUTOS", "AT" & LoopC, CStr(.Stats.UserAtributosBackUP(LoopC)))
        Next LoopC
    End If
    
    For LoopC = 1 To UBound(.Stats.UserSkills)
        Call Manager.ChangeValue("SKILLS", "SK" & LoopC, CStr(.Stats.UserSkills(LoopC)))
        Call Manager.ChangeValue("SKILLS", "ELUSK" & LoopC, CStr(.Stats.EluSkills(LoopC)))
        Call Manager.ChangeValue("SKILLS", "EXPSK" & LoopC, CStr(.Stats.ExpSkills(LoopC)))
    Next LoopC
    
    Call Manager.ChangeValue("CONTACTO", "Email", .email)
        
    Call Manager.ChangeValue("INIT", "Genero", .Genero)
    Call Manager.ChangeValue("INIT", "Raza", .raza)
    Call Manager.ChangeValue("INIT", "Hogar", .Hogar)
    Call Manager.ChangeValue("INIT", "Clase", .clase)
    Call Manager.ChangeValue("INIT", "Desc", .desc)
    
    Call Manager.ChangeValue("INIT", "Heading", CStr(.Char.heading))
    Call Manager.ChangeValue("INIT", "Head", CStr(.OrigChar.Head))

    If .flags.Muerto = 0 Then
        If .Char.Body <> 0 Then
            Call Manager.ChangeValue("INIT", "Body", CStr(.Char.Body))
        End If
    End If
    
    Call Manager.ChangeValue("INIT", "Arma", CStr(.Char.WeaponAnim))
    Call Manager.ChangeValue("INIT", "Escudo", CStr(.Char.ShieldAnim))
    Call Manager.ChangeValue("INIT", "Casco", CStr(.Char.CascoAnim))
    
    'First time around?
    If Manager.GetValue("INIT", "LastIP1") = vbNullString Then
        Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
    'Is it a different ip from last time?
    ElseIf .ip <> Left$(Manager.GetValue("INIT", "LastIP1"), InStr(1, Manager.GetValue("INIT", "LastIP1"), " ") - 1) Then
        Dim i As Integer
        For i = 5 To 2 Step -1
            Call Manager.ChangeValue("INIT", "LastIP" & i, Manager.GetValue("INIT", "LastIP" & CStr(i - 1)))
        Next i
        Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
    'Same ip, just update the date
    Else
        Call Manager.ChangeValue("INIT", "LastIP1", .ip & " - " & Date & ":" & time)
    End If
    
    Call Manager.ChangeValue("INIT", "Position", .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
    
    Call Manager.ChangeValue("STATS", "GLD", CStr(.Stats.GLD))
    Call Manager.ChangeValue("STATS", "BANCO", CStr(.Stats.Banco))
    
    Call Manager.ChangeValue("STATS", "MaxHP", CStr(.Stats.MaxHp))
    Call Manager.ChangeValue("STATS", "MinHP", CStr(.Stats.MinHp))
    
    Call Manager.ChangeValue("STATS", "MaxSTA", CStr(.Stats.MaxSta))
    Call Manager.ChangeValue("STATS", "MinSTA", CStr(.Stats.MinSta))
    
    Call Manager.ChangeValue("STATS", "MaxMAN", CStr(.Stats.MaxMAN))
    Call Manager.ChangeValue("STATS", "MinMAN", CStr(.Stats.MinMAN))
    
    Call Manager.ChangeValue("STATS", "MaxHIT", CStr(.Stats.MaxHIT))
    Call Manager.ChangeValue("STATS", "MinHIT", CStr(.Stats.MinHIT))
    
    Call Manager.ChangeValue("STATS", "MaxAGU", CStr(.Stats.MaxAGU))
    Call Manager.ChangeValue("STATS", "MinAGU", CStr(.Stats.MinAGU))
    
    Call Manager.ChangeValue("STATS", "MaxHAM", CStr(.Stats.MaxHam))
    Call Manager.ChangeValue("STATS", "MinHAM", CStr(.Stats.MinHam))
    
    Call Manager.ChangeValue("STATS", "SkillPtsLibres", CStr(.Stats.SkillPts))
      
    Call Manager.ChangeValue("STATS", "EXP", CStr(.Stats.Exp))
    Call Manager.ChangeValue("STATS", "ELV", CStr(.Stats.ELV))
    
    
    Call Manager.ChangeValue("STATS", "ELU", CStr(.Stats.ELU))
    Call Manager.ChangeValue("MUERTES", "UserMuertes", CStr(.Stats.UsuariosMatados))
    'Call Manager.ChangeValue( "MUERTES", "CrimMuertes", CStr(.Stats.CriminalesMatados))
    Call Manager.ChangeValue("MUERTES", "NpcsMuertes", CStr(.Stats.NPCsMuertos))
      
    '[KEVIN]----------------------------------------------------------------------------
    '*******************************************************************************************
    Call Manager.ChangeValue("BancoInventory", "CantidadItems", val(.BancoInvent.NroItems))
    Dim loopd As Integer
    For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
        Call Manager.ChangeValue("BancoInventory", "Obj" & loopd, .BancoInvent.Object(loopd).ObjIndex & "-" & .BancoInvent.Object(loopd).Amount)
    Next loopd
    '*******************************************************************************************
    '[/KEVIN]-----------
      
    'Save Inv
    Call Manager.ChangeValue("Inventory", "CantidadItems", val(.Invent.NroItems))
    
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call Manager.ChangeValue("Inventory", "Obj" & LoopC, .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).Equipped)
    Next LoopC
    
    Call Manager.ChangeValue("Inventory", "WeaponEqpSlot", CStr(.Invent.WeaponEqpSlot))
    Call Manager.ChangeValue("Inventory", "ArmourEqpSlot", CStr(.Invent.ArmourEqpSlot))
    Call Manager.ChangeValue("Inventory", "CascoEqpSlot", CStr(.Invent.CascoEqpSlot))
    Call Manager.ChangeValue("Inventory", "EscudoEqpSlot", CStr(.Invent.EscudoEqpSlot))
    Call Manager.ChangeValue("Inventory", "BarcoSlot", CStr(.Invent.BarcoSlot))
    Call Manager.ChangeValue("Inventory", "MunicionSlot", CStr(.Invent.MunicionEqpSlot))
    Call Manager.ChangeValue("Inventory", "MochilaSlot", CStr(.Invent.MochilaEqpSlot))
    '/Nacho
    
    Call Manager.ChangeValue("Inventory", "AnilloSlot", CStr(.Invent.AnilloEqpSlot))
    
    'Reputacion
    Call Manager.ChangeValue("REP", "Asesino", CStr(.Reputacion.AsesinoRep))
    Call Manager.ChangeValue("REP", "Bandido", CStr(.Reputacion.BandidoRep))
    Call Manager.ChangeValue("REP", "Burguesia", CStr(.Reputacion.BurguesRep))
    Call Manager.ChangeValue("REP", "Ladrones", CStr(.Reputacion.LadronesRep))
    Call Manager.ChangeValue("REP", "Nobles", CStr(.Reputacion.NobleRep))
    Call Manager.ChangeValue("REP", "Plebe", CStr(.Reputacion.PlebeRep))
    
    Dim L As Long
    L = (-.Reputacion.AsesinoRep) + _
        (-.Reputacion.BandidoRep) + _
        .Reputacion.BurguesRep + _
        (-.Reputacion.LadronesRep) + _
        .Reputacion.NobleRep + _
        .Reputacion.PlebeRep
    L = L / 6
    Call Manager.ChangeValue("REP", "Promedio", CStr(L))
    
    Dim cad As String
    
    For LoopC = 1 To MAXUSERHECHIZOS
        cad = .Stats.UserHechizos(LoopC)
        Call Manager.ChangeValue("HECHIZOS", "H" & LoopC, cad)
    Next
    
    Call SaveQuestStats(UserIndex, Manager) ' GSZAO
    
    Dim NroMascotas As Long
    NroMascotas = .NroMascotas
    
    For LoopC = 1 To MAXMASCOTAS
        ' Mascota valida?
        If .MascotasIndex(LoopC) > 0 Then
            ' Nos aseguramos que la criatura no fue invocada
            If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                cad = .MascotasType(LoopC)
            Else 'Si fue invocada no la guardamos
                cad = "0"
                NroMascotas = NroMascotas - 1
            End If
            Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)
        Else
            cad = .MascotasType(LoopC)
            Call Manager.ChangeValue("MASCOTAS", "MAS" & LoopC, cad)
        End If
    
    Next
    
    Call Manager.ChangeValue("MASCOTAS", "NroMascotas", CStr(NroMascotas))
    
    'Devuelve el head de muerto
    If .flags.Muerto = 1 Then
        .Char.Head = iCabezaMuerto
    End If
End With

Call Manager.DumpFile(UserFile)

Set Manager = Nothing

If Existe Then Call Kill(UserFile & ".bk")

Exit Sub

ErrHandler:
    Call LogError("Error en SaveUser. Error: " & Err.Number & " - " & Err.description & " - Charfile: " & UserFile)
    Set Manager = Nothing

End Sub

Function Criminal(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim L As Long
    
    With UserList(UserIndex).Reputacion
        L = (-.AsesinoRep) + (-.BandidoRep) + .BurguesRep + (-.LadronesRep) + .NobleRep + .PlebeRep
        L = L / 6
        Criminal = (L < 0)
    End With

End Function

Sub BackUPnPc(ByVal NpcIndex As Integer, ByVal hFile As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: ^[GS]^ - 19/05/2012
'10/09/2010 - Pato: Optimice el BackUp de NPCs
'***************************************************

    Dim LoopC As Integer
    
    Print #hFile, "[NPC" & Npclist(NpcIndex).Numero & "]"
    
    With Npclist(NpcIndex)
        'General
        Print #hFile, "Name=" & .Name
        Print #hFile, "ShowName=" & val(.ShowName) ' GSZAO
        Print #hFile, "Desc=" & .desc
        Print #hFile, "Head=" & val(.Char.Head)
        Print #hFile, "Body=" & val(.Char.Body)
        Print #hFile, "Heading=" & val(.Char.heading)
        
        ' GSZAO Tiene escudo, casco o arma?
        If (.Char.ShieldAnim <> 0) Then
            Print #hFile, "ShieldAnim=" & val(.Char.ShieldAnim) ' Escudo
        End If
        If (.Char.CascoAnim <> 0) Then
            Print #hFile, "CascoAnim=" & val(.Char.CascoAnim) ' Casco
        End If
        If (.Char.WeaponAnim <> 0) Then
            Print #hFile, "WeaponAnim=" & val(.Char.WeaponAnim) ' Arma
        End If
        ' GSZAO
        
        Print #hFile, "Movement=" & val(.Movement)
        Print #hFile, "Attackable=" & val(.Attackable)
        Print #hFile, "Comercia=" & val(.Comercia)
        Print #hFile, "TipoItems=" & val(.TipoItems)
        Print #hFile, "Hostil=" & val(.Hostile)
        Print #hFile, "GiveEXP=" & val(.GiveEXP)
        Print #hFile, "GiveGLD=" & val(.GiveGLD)
        Print #hFile, "InvReSpawn=" & val(.InvReSpawn)
        Print #hFile, "NpcType=" & val(.NPCtype)

        'Stats
        Print #hFile, "Alineacion=" & val(.Stats.Alineacion)
        Print #hFile, "DEF=" & val(.Stats.def)
        Print #hFile, "MaxHit=" & val(.Stats.MaxHIT)
        Print #hFile, "MaxHp=" & val(.Stats.MaxHp)
        Print #hFile, "MinHit=" & val(.Stats.MinHIT)
        Print #hFile, "MinHp=" & val(.Stats.MinHp)

        'Flags
        Print #hFile, "ReSpawn=" & val(.flags.Respawn)
        Print #hFile, "BackUp=" & val(.flags.Backup)
        Print #hFile, "Domable=" & val(.flags.Domable)
        
        'Inventario
        Print #hFile, "NroItems=" & val(.Invent.NroItems)
        If .Invent.NroItems > 0 Then
           For LoopC = 1 To .Invent.NroItems
                Print #hFile, "Obj" & LoopC & "=" & .Invent.Object(LoopC).ObjIndex & "-" & .Invent.Object(LoopC).Amount & "-" & .Invent.Object(LoopC).Equipped ' GSZAO
           Next LoopC
        End If
        
        'GSZ Drops
        For LoopC = 1 To MAX_NPC_DROPS
            If .Drop(LoopC).ObjIndex <> 0 Then
                Print #hFile, "Drop" & LoopC & "=" & .Drop(LoopC).ObjIndex & "-" & .Drop(LoopC).Amount & "-" & .Drop(LoopC).Equipped
            End If
        Next LoopC
        ' GSZAO
        
        Print #hFile, ""
    End With

End Sub

Sub CargarNpcBackUp(ByVal NpcIndex As Integer, ByVal NpcNumber As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: ^[GS]^ - 19/05/2012
'
'***************************************************

    'Status
    If frmMain.Visible Then frmMain.txStatus.Text = "Cargando Backup de NPCs"
    
    Dim npcfile As String
    npcfile = pathDats & "NPCs-Backup.dat"

    With Npclist(NpcIndex)
    
        .Numero = NpcNumber
        .Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
        .ShowName = val(GetVar(npcfile, "NPC" & NpcNumber, "ShowName")) ' GSZAO
        
        .desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
        .Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
        .NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))
        
        .Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
        .Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
        .Char.heading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
        
        ' GSZAO Tiene escudo, casco o arma?
        .Char.ShieldAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "ShieldAnim"))
        .Char.CascoAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "CascoAnim"))
        .Char.WeaponAnim = val(GetVar(npcfile, "NPC" & NpcNumber, "WeaponAnim"))
        ' GSZAO
        
        .Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
        .Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
        .Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
        .GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP"))
        
        .GiveGLD = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLD"))
        
        .InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))
        
        .Stats.MaxHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
        .Stats.MinHp = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
        .Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
        .Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
        .Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
        .Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))
        
        Dim LoopC As Integer
        Dim ln As String
        .Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
        If .Invent.NroItems > 0 Then
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
                If (LenB(ln) <> 0) Then ' GSZAO
                    .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                    .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
                    ' GSZAO
                    .Invent.Object(LoopC).Equipped = val(ReadField(3, ln, 45))
                    If .Invent.Object(LoopC).Equipped = 1 Then
                        If ObjData(.Invent.Object(LoopC).ObjIndex).OBJType = otESCUDO Then
                            .Char.ShieldAnim = ObjData(.Invent.Object(LoopC).ObjIndex).ShieldAnim
                        ElseIf ObjData(.Invent.Object(LoopC).ObjIndex).OBJType = otCASCO Then
                            .Char.CascoAnim = ObjData(.Invent.Object(LoopC).ObjIndex).CascoAnim
                        ElseIf ObjData(.Invent.Object(LoopC).ObjIndex).OBJType = otWeapon Then
                            .Char.WeaponAnim = ObjData(.Invent.Object(LoopC).ObjIndex).WeaponAnim
                        End If
                    End If
                    ' GSZAO
                End If
            Next LoopC
        Else
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = 0
                .Invent.Object(LoopC).Amount = 0
            Next LoopC
        End If
        
        For LoopC = 1 To MAX_NPC_DROPS
            ln = GetVar(npcfile, "NPC" & NpcNumber, "Drop" & LoopC)
            If (LenB(ln) <> 0) Then
                .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
                ' GSZAO
                .Drop(LoopC).Equipped = val(ReadField(3, ln, 45))
                If .Drop(LoopC).Equipped = 1 Then
                    If ObjData(.Drop(LoopC).ObjIndex).OBJType = otESCUDO Then
                        .Char.ShieldAnim = ObjData(.Drop(LoopC).ObjIndex).ShieldAnim
                    ElseIf ObjData(.Drop(LoopC).ObjIndex).OBJType = otCASCO Then
                        .Char.CascoAnim = ObjData(.Drop(LoopC).ObjIndex).CascoAnim
                    ElseIf ObjData(.Drop(LoopC).ObjIndex).OBJType = otWeapon Then
                        .Char.WeaponAnim = ObjData(.Drop(LoopC).ObjIndex).WeaponAnim
                    End If
                End If
                ' GSZAO
            End If
        Next LoopC
        
        .flags.NPCActive = True
        .flags.Respawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
        .flags.Backup = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
        .flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
        .flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))
        
        'Tipo de items con los que comercia
        .TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))
    End With

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal Motivo As String)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, UserList(BannedIndex).Name
    Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal Motivo As String)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal Motivo As String)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
    Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer
    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub CargaApuestas()
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Apuestas.Ganancias = val(GetVar(pathDats & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(pathDats & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(pathDats & "apuestas.dat", "Main", "Jugadas"))

End Sub

Public Sub generateMatrix(ByVal mapa As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************

    Dim i As Integer
    Dim j As Integer
    
    ReDim distanceToCities(1 To NumMaps) As HomeDistance
    
    For j = 1 To NUMCIUDADES
        For i = 1 To NumMaps
            distanceToCities(i).distanceToCity(j) = -1
        Next i
    Next j
    
    For j = 1 To NUMCIUDADES
        For i = 1 To 4
            Select Case i
                Case eHeading.NORTH
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.NORTH), j, i, 0, 1)
                Case eHeading.EAST
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.EAST), j, i, 1, 0)
                Case eHeading.SOUTH
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.SOUTH), j, i, 0, 1)
                Case eHeading.WEST
                    Call setDistance(getLimit(Ciudades(j).Map, eHeading.WEST), j, i, -1, 0)
            End Select
        Next i
    Next j

End Sub

Public Sub setDistance(ByVal mapa As Integer, ByVal city As Byte, ByVal side As Integer, Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

Dim i As Integer
Dim lim As Integer

If mapa <= 0 Or mapa > NumMaps Then Exit Sub

If distanceToCities(mapa).distanceToCity(city) >= 0 Then Exit Sub

If mapa = Ciudades(city).Map Then
    distanceToCities(mapa).distanceToCity(city) = 0
Else
    distanceToCities(mapa).distanceToCity(city) = Abs(X) + Abs(Y)
End If

For i = 1 To 4
    lim = getLimit(mapa, i)
    If lim > 0 Then
        Select Case i
            Case eHeading.NORTH
                Call setDistance(lim, city, i, X, Y + 1)
            Case eHeading.EAST
                Call setDistance(lim, city, i, X + 1, Y)
            Case eHeading.SOUTH
                Call setDistance(lim, city, i, X, Y - 1)
            Case eHeading.WEST
                Call setDistance(lim, city, i, X - 1, Y)
        End Select
    End If
Next i
End Sub

Public Function getLimit(ByVal mapa As Integer, ByVal side As Byte) As Integer
'***************************************************
'Author: Budi
'Last Modification: 10/07/2012 - ^[GS]^
'Retrieves the limit in the given side in the given map.
'TODO: This should be set in the .inf map file.
'***************************************************
Dim X As Long
Dim Y As Long

If mapa <= 0 Then Exit Function
If mapa > NumMaps Then Exit Function ' GSZAO

For X = 15 To 87
    For Y = 0 To 3
        Select Case side
            Case eHeading.NORTH
                getLimit = MapData(mapa, X, 7 + Y).TileExit.Map
            Case eHeading.EAST
                getLimit = MapData(mapa, 92 - Y, X).TileExit.Map
            Case eHeading.SOUTH
                getLimit = MapData(mapa, X, 94 - Y).TileExit.Map
            Case eHeading.WEST
                getLimit = MapData(mapa, 9 + Y, X).TileExit.Map
        End Select
        If getLimit > 0 Then Exit Function
    Next Y
Next X
End Function


Public Sub LoadArmadurasFaccion()
'***************************************************
'Author: ZaMa
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************
    Dim ClassIndex As Long
    Dim ArmaduraIndex As Integer
    
    For ClassIndex = 1 To NUMCLASES
    
        ' Defensa minima para armadas altos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para armadas bajos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos altos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        
        ' Defensa minima para caos bajos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMinCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieBaja) = ArmaduraIndex
    
    
        ' Defensa media para armadas altos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para armadas bajos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos altos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        
        ' Defensa media para caos bajos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefMedCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieMedia) = ArmaduraIndex
    
    
        ' Defensa alta para armadas altos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para armadas bajos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaArmyBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Armada(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos altos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosAlto"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Drow).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Elfo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Humano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        
        ' Defensa alta para caos bajos
        ArmaduraIndex = val(GetVar(pathDats & "ArmadurasFaccionarias.dat", "CLASE" & ClassIndex, "DefAltaCaosBajo"))
        
        ArmadurasFaccion(ClassIndex, eRaza.Enano).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
        ArmadurasFaccion(ClassIndex, eRaza.Gnomo).Caos(eTipoDefArmors.ieAlta) = ArmaduraIndex
    
    Next ClassIndex
    
End Sub

