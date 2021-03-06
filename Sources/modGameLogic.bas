Attribute VB_Name = "modExtra"
'Argentum Online 0.12.2
'Copyright (C) 2002 M?rquez Pablo Ignacio
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
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez

Option Explicit

Public Function EsNewbie(ByVal userIndex As Integer) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    EsNewbie = UserList(userIndex).Stats.ELV <= LimiteNewbie
End Function

Public Function esArmada(ByVal userIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************

    esArmada = (UserList(userIndex).fAccion.ArmadaReal = 1)
End Function

Public Function esCaos(ByVal userIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************

    esCaos = (UserList(userIndex).fAccion.FuerzasCaos = 1)
End Function

Public Function EsGm(ByVal userIndex As Integer) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 23/01/2007
'***************************************************

    EsGm = (UserList(userIndex).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero))
End Function

Public Sub DoTileEvents(ByVal userIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 07/09/2012 - ^[GS]^
'Handles the Map passage of Users. Allows the existance
'of exclusive maps for Newbies, Royal Army and Caos Legion members
'and enables GMs to enter every map without restriction.
'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
'***************************************************

    Dim nPos As WorldPos
    Dim FxFlag As Boolean
    Dim TelepRadio As Integer
    Dim GridPos As WorldPos
    Dim DestPos As WorldPos
    
On Error GoTo ErrHandler
    'Controla las salidas
    If InMapBounds(Map, X, Y) Then
        With MapData(Map, X, Y)
            If .ObjInfo.ObjIndex > 0 Then
                FxFlag = ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport
                TelepRadio = ObjData(.ObjInfo.ObjIndex).Radio
            End If
            
            ' WorldGrid GSZAO
            GridPos.Map = 0
            If MaxGrid > 0 Then ' Utiliza Grid
                If X = 10 Then ' izquierda
                    If MapInfo(Map).Grid(eHeading.WEST) <> 0 Then
                        GridPos.Map = MapInfo(Map).Grid(eHeading.WEST)
                        GridPos.X = 89
                        GridPos.Y = Y
                    End If
                ElseIf X = 90 Then ' derecha
                    If MapInfo(Map).Grid(eHeading.EAST) <> 0 Then
                        GridPos.Map = MapInfo(Map).Grid(eHeading.EAST)
                        GridPos.X = 11
                        GridPos.Y = Y
                    End If
                ElseIf Y = 10 Then ' arriba
                    If MapInfo(Map).Grid(eHeading.NORTH) <> 0 Then
                        GridPos.Map = MapInfo(Map).Grid(eHeading.NORTH)
                        GridPos.X = X
                        GridPos.Y = 89
                    End If
                ElseIf Y = 90 Then ' abajo
                    If MapInfo(Map).Grid(eHeading.SOUTH) <> 0 Then
                        GridPos.Map = MapInfo(Map).Grid(eHeading.SOUTH)
                        GridPos.X = X
                        GridPos.Y = 11
                    End If
                End If
            End If
                        
            If (.TileExit.Map > 0 And .TileExit.Map <= NumMaps) Or (GridPos.Map > 0 And GridPos.Map <= NumMaps) Then
                ' Es un teleport, entra en una posicion random, acorde al radio (si es 0, es pos fija)
                ' We have 5 attempts to not falling into another teleport or a map exit.. If we get to the fifth attemp,
                ' the teleport will act as if its radius = 0.
                If FxFlag And TelepRadio > 0 Then
                    Dim attemps As Long
                    Dim exitMap As Boolean
                    Do
                        DestPos.X = .TileExit.X + RandomNumber(TelepRadio * (-1), TelepRadio)
                        DestPos.Y = .TileExit.Y + RandomNumber(TelepRadio * (-1), TelepRadio)
                        
                        attemps = attemps + 1
                        
                        exitMap = MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map > 0 And MapData(.TileExit.Map, DestPos.X, DestPos.Y).TileExit.Map <= NumMaps
                    Loop Until (attemps >= 5 Or exitMap = False)
                    
                    If attemps >= 5 Then
                        DestPos.X = .TileExit.X
                        DestPos.Y = .TileExit.Y
                    End If
                    #If Testeo = 1 Then
                        Call WriteConsoleMsg(userIndex, "TelepRadio de " & Map & "-" & X & "-" & Y & " hacia " & Map & "-" & DestPos.X & "-" & DestPos.Y & ".", FontTypeNames.FONTTYPE_FIGHT)
                    #End If
                
                ' Posicion fija
                ElseIf .TileExit.Map > 0 Then
                    DestPos.Map = .TileExit.Map
                    DestPos.X = .TileExit.X
                    DestPos.Y = .TileExit.Y
                    #If Testeo = 1 Then
                        Call WriteConsoleMsg(userIndex, "TileExit de " & Map & "-" & X & "-" & Y & " hacia " & DestPos.Map & "-" & DestPos.X & "-" & DestPos.Y & ".", FontTypeNames.FONTTYPE_FIGHT)
                    #End If
                ElseIf GridPos.Map > 0 Then ' GSZAO
                    ' NOTA: Los traslados mapeados tienen priodidad al Grid
                    DestPos.Map = GridPos.Map
                    DestPos.X = GridPos.X
                    DestPos.Y = GridPos.Y
                    #If Testeo = 1 Then
                        Call WriteConsoleMsg(userIndex, "GridPos de " & Map & "-" & X & "-" & Y & " hacia " & DestPos.Map & "-" & DestPos.X & "-" & DestPos.Y & ".", FontTypeNames.FONTTYPE_FIGHT)
                    #End If
                End If
                
                If EsGm(userIndex) And .TileExit.Map > 0 Then ' 0.13.3
                    ' Solo generamos el Log cuando teletransporte mediante un telep normal y no con el paso de mapa por grid
                    Call LogGM(UserList(userIndex).name, "Utiliz? un teleport hacia el mapa " & _
                        DestPos.Map & " (" & DestPos.X & "," & DestPos.Y & ")")
                End If
                
                ' Si es un mapa que no admite muertos
                If MapInfo(DestPos.Map).OnDeathGoTo.Map <> 0 Then ' 0.13.3
                    ' Si esta muerto no puede entrar
                    If UserList(userIndex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(userIndex, "S?lo se permite entrar al mapa a los personajes vivos.", FontTypeNames.FONTTYPE_INFO)
                        Call ClosestStablePos(UserList(userIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                        
                        Exit Sub
                    End If
                End If
                
                '?Es mapa de newbies?
                If MapInfo(DestPos.Map).Restringir = eRestrict.restrict_newbie Then
                    '?El usuario es un newbie?
                    If EsNewbie(userIndex) Or EsGm(userIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(userIndex)) Then
                            Call WarpUserChar(userIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es newbie
                        Call WriteMensajes(userIndex, eMensajes.Mensaje027) '"Mapa exclusivo para newbies."
                        Call ClosestStablePos(UserList(userIndex).Pos, nPos)
        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, False)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_armada Then '?Es mapa de Armadas?
                    '?El usuario es Armada?
                    If esArmada(userIndex) Or EsGm(userIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(userIndex)) Then
                            Call WarpUserChar(userIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es armada
                        Call WriteMensajes(userIndex, eMensajes.Mensaje028) '"Mapa exclusivo para miembros del ej?rcito real."
                        Call ClosestStablePos(UserList(userIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_caos Then '?Es mapa de Caos?
                    '?El usuario es Caos?
                    If esCaos(userIndex) Or EsGm(userIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(userIndex)) Then
                            Call WarpUserChar(userIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es caos
                        Call WriteMensajes(userIndex, eMensajes.Mensaje029) '"Mapa exclusivo para miembros de la legi?n oscura."
                        Call ClosestStablePos(UserList(userIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                ElseIf MapInfo(DestPos.Map).Restringir = eRestrict.restrict_faccion Then '?Es mapa de faccionarios?
                    '?El usuario es Armada o Caos?
                    If esArmada(userIndex) Or esCaos(userIndex) Or EsGm(userIndex) Then
                        If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(userIndex)) Then
                            Call WarpUserChar(userIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                        Else
                            Call ClosestLegalPos(DestPos, nPos)
                            If nPos.X <> 0 And nPos.Y <> 0 Then
                                Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                            End If
                        End If
                    Else 'No es Faccionario
                        Call WriteMensajes(userIndex, eMensajes.Mensaje030) '"Solo se permite entrar al mapa si eres miembro de alguna facci?n."
                        Call ClosestStablePos(UserList(userIndex).Pos, nPos)
                        
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                Else 'No es un mapa de newbies, ni Armadas, ni Caos, ni faccionario.
                    If LegalPos(DestPos.Map, DestPos.X, DestPos.Y, PuedeAtravesarAgua(userIndex)) Then
                        Call WarpUserChar(userIndex, DestPos.Map, DestPos.X, DestPos.Y, FxFlag)
                    Else
                        Call ClosestLegalPos(DestPos, nPos)
                        If nPos.X <> 0 And nPos.Y <> 0 Then
                            Call WarpUserChar(userIndex, nPos.Map, nPos.X, nPos.Y, FxFlag)
                        End If
                    End If
                End If
                
                'Te fusite del mapa. La criatura ya no es m?s tuya ni te reconoce como que vos la atacaste.
                Dim aN As Integer
                
                aN = UserList(userIndex).flags.AtacadoPorNpc
                If aN > 0 Then
                   Npclist(aN).Movement = Npclist(aN).flags.OldMovement
                   Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
                   Npclist(aN).flags.AttackedBy = vbNullString
                End If
            
                aN = UserList(userIndex).flags.NPCAtacado
                If aN > 0 Then
                    If Npclist(aN).flags.AttackedFirstBy = UserList(userIndex).name Then
                        Npclist(aN).flags.AttackedFirstBy = vbNullString
                    End If
                End If
                UserList(userIndex).flags.AtacadoPorNpc = 0
                UserList(userIndex).flags.NPCAtacado = 0
            End If
        
        End With
    End If
Exit Sub

ErrHandler:
    Call LogError("DoTileEvents::Error " & Err.Number & " - Desc: " & Err.description & " - Userindex: " & userIndex)
End Sub

Public Function InVisionRangeAndMap(ByVal userIndex As Integer, ByRef OtherUserPos As WorldPos) As Boolean ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************
    
    With UserList(userIndex)
        ' Same map?
        If .Pos.Map <> OtherUserPos.Map Then Exit Function
        ' In x range?
        If OtherUserPos.X < .Pos.X - MinXBorder Or OtherUserPos.X > .Pos.X + MinXBorder Then Exit Function
        ' In y range?
        If OtherUserPos.Y < .Pos.Y - MinYBorder And OtherUserPos.Y > .Pos.Y + MinYBorder Then Exit Function
    End With

    InVisionRangeAndMap = True
    
End Function

Function InRangoVision(ByVal userIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    If X > UserList(userIndex).Pos.X - MinXBorder And X < UserList(userIndex).Pos.X + MinXBorder Then
        If Y > UserList(userIndex).Pos.Y - MinYBorder And Y < UserList(userIndex).Pos.Y + MinYBorder Then
            InRangoVision = True
            Exit Function
        End If
    End If
    InRangoVision = False

End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, X As Integer, Y As Integer) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    If X > Npclist(NpcIndex).Pos.X - MinXBorder And X < Npclist(NpcIndex).Pos.X + MinXBorder Then
        If Y > Npclist(NpcIndex).Pos.Y - MinYBorder And Y < Npclist(NpcIndex).Pos.Y + MinYBorder Then
            InRangoVisionNPC = True
            Exit Function
        End If
    End If
    InRangoVisionNPC = False

End Function


Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        InMapBounds = False
    Else
        InMapBounds = True
    End If
    
End Function
    
Private Function RhombLegalPos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                               ByVal Distance As Long, Optional PuedeAgua As Boolean = False, _
                               Optional PuedeTierra As Boolean = True, _
                               Optional ByVal CheckExitTile As Boolean = False) As Boolean ' 0.13.3
'***************************************************
'Author: Marco Vanotti (Marco)
'Last Modification: 10/07/2012 - ^[GS]^
' walks all the perimeter of a rhomb of side  "distance + 1",
' which starts at Pos.x - Distance and Pos.y
'***************************************************

    Dim i As Long
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX + i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX + i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX + i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX - i, vY + i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY + i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1
        If (LegalPos(Pos.Map, vX - i, vY - i, PuedeAgua, PuedeTierra, CheckExitTile)) Then
            vX = vX - i
            vY = vY - i
            RhombLegalPos = True
            Exit Function
        End If
    Next
    
    RhombLegalPos = False
    
End Function

Public Function RhombLegalTilePos(ByRef Pos As WorldPos, ByRef vX As Long, ByRef vY As Long, _
                                  ByVal Distance As Long, ByVal ObjIndex As Integer, ByVal ObjAmount As Long, _
                                  ByVal PuedeAgua As Boolean, ByVal PuedeTierra As Boolean) As Boolean ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 10/07/2012 - ^[GS]^
' walks all the perimeter of a rhomb of side  "distance + 1",
' which starts at Pos.x - Distance and Pos.y
' and searchs for a valid position to drop items
'***************************************************
On Error GoTo ErrHandler

    Dim i As Long
    Dim HayObj As Boolean
    
    Dim X As Integer
    Dim Y As Integer
    Dim MapObjIndex As Integer
    
    vX = Pos.X - Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY - i
        
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
            
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y - Distance
    
    For i = 0 To Distance - 1
        
        X = vX + i
        Y = vY + i
        
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X + Distance
    vY = Pos.Y
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY + i
    
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
        
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    vX = Pos.X
    vY = Pos.Y + Distance
    
    For i = 0 To Distance - 1
        
        X = vX - i
        Y = vY - i
    
        If (LegalPos(Pos.Map, X, Y, PuedeAgua, PuedeTierra, True)) Then
            ' No hay obj tirado o la suma de lo que hay + lo nuevo <= 10k
            If Not HayObjeto(Pos.Map, X, Y, ObjIndex, ObjAmount) Then
                vX = X
                vY = Y
                
                RhombLegalTilePos = True
                Exit Function
            End If
        End If
    Next
    
    RhombLegalTilePos = False
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en RhombLegalTilePos. Error: " & Err.Number & " - " & Err.description)
End Function

Public Function HayObjeto(ByVal mapa As Integer, ByVal X As Long, ByVal Y As Long, _
                          ByVal ObjIndex As Integer, ByVal ObjAmount As Long) As Boolean ' 0.13.3
'***************************************************
'Author: ZaMa
'Last Modification: 10/07/2012 - ^[GS]^
'Checks if there's space in a tile to add an itemAmount
'***************************************************
    Dim MapObjIndex As Integer
    MapObjIndex = MapData(mapa, X, Y).ObjInfo.ObjIndex
            
    ' Hay un objeto tirado?
    If MapObjIndex <> 0 Then
        ' Es el mismo objeto?
        If MapObjIndex = ObjIndex Then
            ' La suma es menor a 10k?
            HayObjeto = (MapData(mapa, X, Y).ObjInfo.Amount + ObjAmount > MAX_INVENTORY_OBJS)
        Else
            HayObjeto = True
        End If
    Else
        HayObjeto = False
    End If

End Function


Sub ClosestLegalPos(Pos As WorldPos, ByRef nPos As WorldPos, Optional PuedeAgua As Boolean = False, _
                    Optional PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False)
'*****************************************************************
'Author: Unknownn (original version)
'Last Modification: 10/07/2012 - ^[GS]^
'History:
' - 01/24/2007 (ToxicWaste)
'Encuentra la posicion legal mas cercana y la guarda en nPos
'*****************************************************************

    Dim Found As Boolean
    Dim loopC As Integer
    Dim tX As Long
    Dim tY As Long

    nPos = Pos
    tX = Pos.X
    tY = Pos.Y
    
    loopC = 1

    ' La primera posicion es valida?
    If LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, CheckExitTile) Then
        Found = True

    ' Busca en las demas posiciones, en forma de "rombo"
    Else
        While (Not Found) And loopC <= 12
            If RhombLegalPos(Pos, tX, tY, loopC, PuedeAgua, PuedeTierra, CheckExitTile) Then
                nPos.X = tX
                nPos.Y = tY
                Found = True
            End If
        
            loopC = loopC + 1
        Wend
    End If
    
    If Not Found Then
        nPos.X = 0
        nPos.Y = 0
    End If

End Sub

Private Sub ClosestStablePos(Pos As WorldPos, ByRef nPos As WorldPos)
'***************************************************
'Author: Unknownn
'Last Modification: 10/07/2012 - ^[GS]^
'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
'*****************************************************************

Call ClosestLegalPos(Pos, nPos, , , True)

End Sub

Function NameIndex(ByVal name As String) As Integer
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim userIndex As Long
    
    '?Nombre valido?
    If LenB(name) = 0 Then
        NameIndex = 0
        Exit Function
    End If
    
    If InStrB(name, "+") <> 0 Then
        name = UCase$(Replace(name, "+", " "))
    End If
    
    userIndex = 1
    Do Until UCase$(UserList(userIndex).name) = UCase$(name)
        
        userIndex = userIndex + 1
        
        If userIndex > iniMaxUsuarios Then
            NameIndex = 0
            Exit Function
        End If
    Loop
     
    NameIndex = userIndex
End Function

Function CheckForSameIP(ByVal userIndex As Integer, ByVal UserIP As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    Dim loopC As Long
    
    For loopC = 1 To iniMaxUsuarios
        If UserList(loopC).flags.UserLogged = True Then
            If UserList(loopC).ip = UserIP And userIndex <> loopC Then
                CheckForSameIP = True
                Exit Function
            End If
        End If
    Next loopC
    
    CheckForSameIP = False
End Function

Function CheckForSameName(ByVal name As String) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

'Controlo que no existan usuarios con el mismo nombre
    Dim loopC As Long
    
    For loopC = 1 To LastUser
        If UserList(loopC).flags.UserLogged Then
            
            'If UCase$(UserList(LoopC).Name) = UCase$(Name) And UserList(LoopC).ConnID <> -1 Then
            'OJO PREGUNTAR POR EL CONNID <> -1 PRODUCE QUE UN PJ EN DETERMINADO
            'MOMENTO PUEDA ESTAR LOGUEADO 2 VECES (IE: CIERRA EL SOCKET DESDE ALLA)
            'ESE EVENTO NO DISPARA UN SAVE USER, LO QUE PUEDE SER UTILIZADO PARA DUPLICAR ITEMS
            'ESTE BUG EN ALKON PRODUJO QUE EL SERVIDOR ESTE CAIDO DURANTE 3 DIAS. ATENTOS.
            
            If UCase$(UserList(loopC).name) = UCase$(name) Then
                CheckForSameName = True
                Exit Function
            End If
        End If
    Next loopC
    
    CheckForSameName = False
End Function

Sub HeadtoPos(ByVal Head As eHeading, ByRef Pos As WorldPos)
'***************************************************
'Author: Unknownn
'Last Modification: -
'Toma una posicion y se mueve hacia donde esta perfilado
'*****************************************************************

    Select Case Head
        Case eHeading.NORTH
            Pos.Y = Pos.Y - 1
        
        Case eHeading.SOUTH
            Pos.Y = Pos.Y + 1
        
        Case eHeading.EAST
            Pos.X = Pos.X + 1
        
        Case eHeading.WEST
            Pos.X = Pos.X - 1
    End Select
End Sub

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal CheckExitTile As Boolean = False) As Boolean
'***************************************************
'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
'Last Modification: 27/12/2012 - ^[GS]^
'Checks if the position is Legal.
'***************************************************

    '?Es un mapa valido?
    If (Map <= 0 Or Map > NumMaps) Or _
       (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPos = False
    Else
        With MapData(Map, X, Y)
            If PuedeAgua And PuedeTierra Then
                LegalPos = (.Blocked <> 1) And _
                           (.userIndex = 0) And _
                           (.NpcIndex = 0)
            ElseIf PuedeTierra And Not PuedeAgua Then
                LegalPos = (.Blocked <> 1) And _
                           (.userIndex = 0) And _
                           (.NpcIndex = 0) And _
                           (Not HayAgua(Map, X, Y))
            ElseIf PuedeAgua And Not PuedeTierra Then
                LegalPos = (.Blocked <> 1) And _
                           (.userIndex = 0) And _
                           (.NpcIndex = 0) And _
                           (HayAgua(Map, X, Y))
            ElseIf ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otDestruible Then ' GSZAO
                LegalPos = False
            Else
                LegalPos = False
            End If
        End With
        
        If CheckExitTile Then
            LegalPos = LegalPos And (MapData(Map, X, Y).TileExit.Map = 0)
        End If
    End If


End Function

Function MoveToLegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 27/12/2012 - ^[GS]^
'Checks if the position is Legal, but considers that if there's a casper, it's a legal movement.
'***************************************************

Dim userIndex As Integer
Dim IsDeadChar As Boolean
Dim IsAdminInvisible As Boolean


'?Es un mapa valido?
If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        MoveToLegalPos = False
Else
    With MapData(Map, X, Y)
        userIndex = .userIndex
        
        If userIndex > 0 Then
            IsDeadChar = (UserList(userIndex).flags.Muerto = 1)
            IsAdminInvisible = (UserList(userIndex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
        
        If PuedeAgua And PuedeTierra Then
            MoveToLegalPos = (.Blocked <> 1) And (userIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0)
        ElseIf PuedeTierra And Not PuedeAgua Then
            MoveToLegalPos = (.Blocked <> 1) And (userIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (Not HayAgua(Map, X, Y))
        ElseIf PuedeAgua And Not PuedeTierra Then
            MoveToLegalPos = (.Blocked <> 1) And (userIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (HayAgua(Map, X, Y))
        ElseIf ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otDestruible Then ' GSZAO
            MoveToLegalPos = False
        Else
            MoveToLegalPos = False
        End If
    End With
End If

End Function

Public Sub FindLegalPos(ByVal userIndex As Integer, ByVal Map As Integer, ByRef X As Integer, ByRef Y As Integer)
'***************************************************
'Autor: ZaMa
'Last Modification: 26/12/2012 - ^[GS]^
'Search for a Legal pos for the user who is being teleported.
'***************************************************

    If MapData(Map, X, Y).userIndex <> 0 Or MapData(Map, X, Y).NpcIndex <> 0 Then
                    
        ' Se teletransporta a la misma pos a la que estaba
        If MapData(Map, X, Y).userIndex = userIndex Then Exit Sub
                            
        Dim FoundPlace As Boolean
        Dim tX As Long
        Dim tY As Long
        Dim Rango As Long
        Dim OtherUserIndex As Integer
    
        For Rango = 1 To 5
            For tY = Y - Rango To Y + Rango
                For tX = X - Rango To X + Rango
                    'Reviso que no haya User ni NPC
                    If MapData(Map, tX, tY).userIndex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        If ObjData(MapData(Map, tX, tY).ObjInfo.ObjIndex).OBJType <> eOBJType.otDestruible Then ' GSZAO
                            If InMapBounds(Map, tX, tY) Then FoundPlace = True
                            Exit For
                        End If
                    End If

                Next tX
        
                If FoundPlace Then Exit For
            Next tY
            
            If FoundPlace Then Exit For
        Next Rango

    
        If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
            X = tX
            Y = tY
        Else
            'Muy poco probable, pero..
            'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
            OtherUserIndex = MapData(Map, X, Y).userIndex
            If OtherUserIndex <> 0 Then
                'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
                If UserList(OtherUserIndex).ComUsu.DestUsu > 0 Then
                    'Le avisamos al que estaba comerciando que se tuvo que ir.
                    If UserList(UserList(OtherUserIndex).ComUsu.DestUsu).flags.UserLogged Then
                        Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu)
                        Call WriteMensajes(UserList(OtherUserIndex).ComUsu.DestUsu, eMensajes.Mensaje031) '"Comercio cancelado. El otro usuario se ha desconectado."
                        Call FlushBuffer(UserList(OtherUserIndex).ComUsu.DestUsu)
                    End If
                    'Lo sacamos.
                    If UserList(OtherUserIndex).flags.UserLogged Then
                        Call FinComerciarUsu(OtherUserIndex)
                        Call WriteErrorMsg(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor recon?ctate...")
                        Call FlushBuffer(OtherUserIndex)
                    End If
                End If
            
                Call CloseSocket(OtherUserIndex)
            End If
        End If
    End If

End Sub

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
'***************************************************
'Autor: Unkwnown
'Last Modification: 01/09/2013 - ^[GS]^
'Checks if it's a Legal pos for the npc to move to.
'***************************************************
    Dim IsDeadChar As Boolean
    Dim userIndex As Integer
    Dim IsAdminInvisible As Boolean
    
    If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        LegalPosNPC = False
        Exit Function
    End If

    With MapData(Map, X, Y)
        userIndex = .userIndex
        If userIndex > 0 Then
            IsDeadChar = UserList(userIndex).flags.Muerto = 1
            IsAdminInvisible = (UserList(userIndex).flags.AdminInvisible = 1)
        Else
            IsDeadChar = False
            IsAdminInvisible = False
        End If
    
        If AguaValida = 0 Then
            LegalPosNPC = (.Blocked <> 1) And (.userIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.SINNPCS Or .trigger <> eTrigger.BAJOTECHOSINNPCS Or IsPet) And Not HayAgua(Map, X, Y)
        Else
            LegalPosNPC = (.Blocked <> 1) And (.userIndex = 0 Or IsDeadChar Or IsAdminInvisible) And (.NpcIndex = 0) And (.trigger <> eTrigger.SINNPCS Or .trigger <> eTrigger.BAJOTECHOSINNPCS Or IsPet)
        End If
    End With
End Function

Sub SendHelp(ByVal Index As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

Dim NumHelpLines As Integer
Dim loopC As Integer

NumHelpLines = val(GetVar(pathDats & "Help.dat", "INIT", "NumLines"))

For loopC = 1 To NumHelpLines
    Call WriteConsoleMsg(Index, GetVar(pathDats & "Help.dat", "Help", "Line" & loopC), FontTypeNames.FONTTYPE_INFO)
Next loopC

End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal userIndex As Integer)
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    If Npclist(NpcIndex).NroExpresiones > 0 Then
        Dim randomi
        randomi = RandomNumber(1, Npclist(NpcIndex).NroExpresiones)
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageChatOverHead(Npclist(NpcIndex).Expresiones(randomi), Npclist(NpcIndex).Char.CharIndex, vbWhite))
    End If
    
End Sub

Sub LookatTile(ByVal userIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 09/06/2013 - ^[GS]^
'
'***************************************************
On Error GoTo ErrHandler

'Responde al click del usuario sobre el mapa
Dim FoundChar As Byte
Dim FoundSomething As Byte
Dim TempCharIndex As Integer
Dim Stat As String
Dim ft As FontTypeNames

With UserList(userIndex)
    '?Rango Visi?n? (ToxicWaste)
    If (Abs(.Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(.Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '?Posicion valida?
    If InMapBounds(Map, X, Y) Then
        With .flags
            .TargetMap = Map
            .TargetX = X
            .TargetY = Y
            '?Es un obj?
            If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                'Informa el nombre
                .TargetObjMap = Map
                .TargetObjX = X
                .TargetObjY = Y
                FoundSomething = 1
            ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
                'Informa el nombre
                If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                    .TargetObjMap = Map
                    .TargetObjX = X + 1
                    .TargetObjY = Y
                    FoundSomething = 1
                End If
            ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
                If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = X + 1
                    .TargetObjY = Y + 1
                    FoundSomething = 1
                End If
            ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
                If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = eOBJType.otPuertas Then
                    'Informa el nombre
                    .TargetObjMap = Map
                    .TargetObjX = X
                    .TargetObjY = Y + 1
                    FoundSomething = 1
                End If
            End If
            
            If FoundSomething = 1 Then
                .targetObj = MapData(Map, .TargetObjX, .TargetObjY).ObjInfo.ObjIndex
                If ObjData(MapData(Map, .TargetObjX, .TargetObjY).ObjInfo.ObjIndex).OBJType = otDestruible Then ' GSZAO
                    Call WriteConsoleMsg(userIndex, ObjData(.targetObj).name & " (Resistencia: " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.ExtraLong & ")", FontTypeNames.FONTTYPE_OBJ)
                Else
                    If MostrarCantidad(.targetObj) And ObjData(.targetObj).NoAgarrable = 0 Then
                        If .targetObj = iORO Then ' GSZAO  ?Es oro?
                            Call WriteConsoleMsg(userIndex, ObjData(.targetObj).name & " (Cant.: " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & ")", FontTypeNames.FONTTYPE_GOLD)
                        Else
                            Call WriteConsoleMsg(userIndex, ObjData(.targetObj).name & " (Cant.: " & MapData(.TargetObjMap, .TargetObjX, .TargetObjY).ObjInfo.Amount & ")", FontTypeNames.FONTTYPE_OBJ)
                        End If
                    Else
                        Call WriteConsoleMsg(userIndex, ObjData(.targetObj).name, FontTypeNames.FONTTYPE_OBJ)
                    End If
                End If
                
            End If
            '?Es un personaje?
            If Y + 1 <= YMaxMapSize Then
                If MapData(Map, X, Y + 1).userIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y + 1).userIndex
                    FoundChar = 1
                End If
                If MapData(Map, X, Y + 1).NpcIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
                    FoundChar = 2
                End If
            End If
            '?Es un personaje?
            If FoundChar = 0 Then
                If MapData(Map, X, Y).userIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y).userIndex
                    FoundChar = 1
                End If
                If MapData(Map, X, Y).NpcIndex > 0 Then
                    TempCharIndex = MapData(Map, X, Y).NpcIndex
                    FoundChar = 2
                End If
            End If
        End With
    
    
        'Reaccion al personaje
        If FoundChar = 1 Then '  ?Encontro un Usuario?
           If UserList(TempCharIndex).flags.AdminInvisible = 0 Or .flags.Privilegios And PlayerType.Dios Then
                With UserList(TempCharIndex)
                    If LenB(.DescRM) = 0 And .ShowName Then 'No tiene descRM y quiere que se vea su nombre.
                        If EsNewbie(TempCharIndex) Then
                            Stat = " <NEWBIE>"
                        End If
                        
                        If .fAccion.ArmadaReal = 1 Then
                            Stat = Stat & " <Ej?rcito Real> " & "<" & TituloReal(TempCharIndex) & ">"
                        ElseIf .fAccion.FuerzasCaos = 1 Then
                            Stat = Stat & " <Legi?n Oscura> " & "<" & TituloCaos(TempCharIndex) & ">"
                        End If
                        
                        If .GuildIndex > 0 Then
                            Stat = Stat & " <" & modGuilds.GuildName(.GuildIndex) & ">"
                        End If
                        
                        If Len(.desc) > 0 Then
                            Stat = "Ves a " & .name & Stat & " - " & .desc
                        Else
                            Stat = "Ves a " & .name & Stat
                        End If
                        
                                        
                        If .flags.Privilegios And PlayerType.RoyalCouncil Then
                            Stat = Stat & " [CONSEJO DE BANDERBILL]"
                            ft = FontTypeNames.FONTTYPE_CONSEJOVesA
                        ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                            Stat = Stat & " [CONCILIO DE LAS SOMBRAS]"
                            ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                        Else
                            If Not .flags.Privilegios And PlayerType.User Then
                                If .flags.Privilegios = PlayerType.Dios Then
                                    Stat = Stat & " <DIOS>"
                                    ft = FontTypeNames.FONTTYPE_DIOS
                                ElseIf .flags.Privilegios = PlayerType.SemiDios Then
                                    Stat = Stat & " <SEMIDIOS>"
                                    ft = FontTypeNames.FONTTYPE_SEMIDIOS
                                ElseIf .flags.Privilegios = PlayerType.Consejero Then
                                    Stat = Stat & " <CONSEJERO>"
                                    ft = FontTypeNames.FOTNTYPE_CONSEJERO
                                ElseIf .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Consejero) Or .flags.Privilegios = (PlayerType.RoleMaster Or PlayerType.Dios) Then
                                    Stat = Stat & " <ROLMASTER>"
                                    ft = FontTypeNames.FONTTYPE_EJECUCION
                                Else ' Rango Admin :P
                                    Stat = Stat & " <ADMIN>"
                                    ft = FontTypeNames.FONTTYPE_EJECUCION
                                End If
                                
                            ElseIf Criminal(TempCharIndex) Then
                                Stat = Stat & " <CRIMINAL>"
                                ft = FontTypeNames.FONTTYPE_FIGHT
                            Else
                                Stat = Stat & " <CIUDADANO>"
                                ft = FontTypeNames.FONTTYPE_CITIZEN
                            End If
                        End If
                        
                        If LenB(.flags.Matrimonio) > 0 Then ' GSZAO
                            If .Genero = eGenero.Mujer Then
                                Stat = Stat & " [CASADA CON " & .flags.Matrimonio & "]"
                            Else
                                Stat = Stat & " [CASADO CON " & .flags.Matrimonio & "]"
                            End If
                        End If
                        
                        If .flags.Muerto = 1 Then ' GSZAO Stat de Muerto
                            Stat = Stat & " [MUERTO]"
                            ft = FontTypeNames.FONTTYPE_EJECUCION
                        End If
                        
                    Else  'Si tiene descRM la muestro siempre.
                        Stat = .DescRM
                        ft = FontTypeNames.FONTTYPE_INFOBOLD
                    End If
                End With
                
                If LenB(Stat) > 0 Then
                    Call WriteConsoleMsg(userIndex, Stat, ft)
                End If
                
                FoundSomething = 1
                .flags.targetUser = TempCharIndex
                .flags.TargetNPC = 0
                .flags.TargetNpcTipo = eNPCType.Comun
           End If
        End If
    
        With .flags
            If FoundChar = 2 Then '?Encontro un NPC?
                Dim estatus As String
                Dim MinHp As Long
                Dim MaxHp As Long
                Dim SupervivenciaSkill As Byte
                Dim sDesc As String
                
                MinHp = Npclist(TempCharIndex).Stats.MinHp
                MaxHp = Npclist(TempCharIndex).Stats.MaxHp
                SupervivenciaSkill = UserList(userIndex).Stats.UserSkills(eSkill.Supervivencia)
                
                If .Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
                    estatus = "(" & MinHp & "/" & MaxHp & ")"
                Else
                    If .Muerto = 0 Then
                        If SupervivenciaSkill <= 10 Then
                            estatus = "(Dudoso)"
                        ElseIf SupervivenciaSkill <= 20 Then
                            If MinHp < (MaxHp / 2) Then
                                estatus = "(Malherido)"
                            Else
                                estatus = "(Dudoso)"
                            End If
                        ElseIf SupervivenciaSkill <= 30 Then
                            If MinHp < (MaxHp * 0.5) Then
                                estatus = "(Malherido)"
                            ElseIf MinHp < (MaxHp * 0.75) Then
                                estatus = "(Herido)"
                            Else
                                estatus = "(Dudoso)"
                            End If
                        ElseIf SupervivenciaSkill <= 40 Then
                            If MinHp < (MaxHp * 0.25) Then
                                estatus = "(Muy malherido)"
                            ElseIf MinHp < (MaxHp * 0.5) Then
                                estatus = "(Malherido)"
                            ElseIf MinHp < (MaxHp * 0.75) Then
                                estatus = "(Herido)"
                            Else
                                estatus = "(Dudoso)"
                            End If
                        ElseIf SupervivenciaSkill < 60 Then
                            If MinHp < (MaxHp * 0.05) Then
                                estatus = "(Agonizando)"
                            ElseIf MinHp < (MaxHp * 0.1) Then
                                estatus = "(Casi muerto)"
                            ElseIf MinHp < (MaxHp * 0.25) Then
                                estatus = "(Muy Malherido)"
                            ElseIf MinHp < (MaxHp * 0.5) Then
                                estatus = "(Malherido)"
                            ElseIf MinHp < (MaxHp * 0.75) Then
                                estatus = "(Herido)"
                            ElseIf MinHp < (MaxHp) Then
                                estatus = "(Levemente herido)"
                            Else
                                estatus = "(Intacto)"
                            End If
                        Else
                            estatus = "(" & MinHp & "/" & MaxHp & ")"
                        End If
                    End If
                End If
                
                If Len(Npclist(TempCharIndex).desc) > 1 Then
                    Stat = Npclist(TempCharIndex).desc
                    
                    '?Es el rey o el demonio?
                    If Npclist(TempCharIndex).NPCtype = eNPCType.Noble Then
                        If Npclist(TempCharIndex).flags.fAccion = 0 Then 'Es el Rey.
                            'Si es de la Legi?n Oscura y usuario com?n mostramos el mensaje correspondiente y lo ejecutamos:
                            If UserList(userIndex).fAccion.FuerzasCaos = 1 Then
                                Stat = MENSAJE_REY_CAOS
                                If .Privilegios And PlayerType.User Then
                                    If .Muerto = 0 Then Call UserDie(userIndex)
                                End If
                            ElseIf Criminal(userIndex) Then
                            'Nos fijamos si es criminal enlistable o no enlistable:
                                If UserList(userIndex).fAccion.CiudadanosMatados > 0 Or _
                                UserList(userIndex).fAccion.Reenlistadas > 4 Then 'Es criminal no enlistable.
                                    Stat = MENSAJE_REY_CRIMINAL_NOENLISTABLE
                                Else 'Es criminal enlistable.
                                    Stat = MENSAJE_REY_CRIMINAL_ENLISTABLE
                                End If
                            End If
                        Else 'Es el demonio
                            'Si es de la Armada Real y usuario com?n mostramos el mensaje correspondiente y lo ejecutamos:
                            If UserList(userIndex).fAccion.ArmadaReal = 1 Then
                                Stat = MENSAJE_DEMONIO_REAL
                                '
                                If .Privilegios And PlayerType.User Then
                                    If .Muerto = 0 Then Call UserDie(userIndex)
                                End If
                            ElseIf Not Criminal(userIndex) Then
                            'Nos fijamos si es ciudadano enlistable o no enlistable:
                                If UserList(userIndex).fAccion.RecibioExpInicialReal = 1 Or _
                                UserList(userIndex).fAccion.Reenlistadas > 4 Then 'Es ciudadano no enlistable.
                                    Stat = MENSAJE_DEMONIO_CIUDADANO_NOENLISTABLE
                                Else 'Es ciudadano enlistable.
                                    Stat = MENSAJE_DEMONIO_CIUDADANO_ENLISTABLE
                                End If
                            End If
                        End If
                    End If
                    
                    'Enviamos el mensaje propiamente dicho:
                    Call WriteChatOverHead(userIndex, Stat, Npclist(TempCharIndex).Char.CharIndex, vbWhite)
                Else
                
                    Dim CentinelaIndex As Integer
                    CentinelaIndex = EsCentinela(TempCharIndex)
                    
                    If CentinelaIndex <> 0 Then
                        'Enviamos nuevamente el texto del centinela seg?n quien pregunta
                        Call modCentinela.CentinelaSendClave(userIndex, CentinelaIndex)
                    Else
                        If Npclist(TempCharIndex).MaestroUser > 0 Then
                            If Npclist(TempCharIndex).MaestroUser = userIndex Then ' NPC propio
                                Call WriteConsoleMsg(userIndex, Npclist(TempCharIndex).name & " " & estatus & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name & ".", FontTypeNames.FONTTYPE_NPC_PEACE)
                            Else ' NPC de otro usuario...
                                Call WriteConsoleMsg(userIndex, Npclist(TempCharIndex).name & " " & estatus & " es mascota de " & UserList(Npclist(TempCharIndex).MaestroUser).name & ".", FontTypeNames.FONTTYPE_NPC_WARNING)
                            End If
                        Else
                            sDesc = Npclist(TempCharIndex).name & " " & estatus
                            If Npclist(TempCharIndex).Owner > 0 Then sDesc = sDesc & " le pertenece a " & UserList(Npclist(TempCharIndex).Owner).name
                            sDesc = sDesc & "."
                            
                            If Npclist(TempCharIndex).Hostile = 0 Then ' NPC no hostil
                                Call WriteConsoleMsg(userIndex, sDesc, FontTypeNames.FONTTYPE_NPC_PEACE)
                            Else ' NPC  hostil
                                Call WriteConsoleMsg(userIndex, sDesc, FontTypeNames.FONTTYPE_NPC_WARNING)
                            End If
                            
                            If .Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                                If (Npclist(TempCharIndex).flags.AttackedFirstBy <> vbNullString) Then ' solo si alguien le peg?!
                                    Call WriteConsoleMsg(userIndex, "Le peg? primero: " & Npclist(TempCharIndex).flags.AttackedFirstBy & ".", FontTypeNames.FONTTYPE_INFO)
                                End If
                            End If
                        End If
                    End If
                End If
                
                FoundSomething = 1
                .TargetNpcTipo = Npclist(TempCharIndex).NPCtype
                .TargetNPC = TempCharIndex
                .targetUser = 0
                .targetObj = 0
            End If
            
            If FoundChar = 0 Then
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .targetUser = 0
            End If
            
            '*** NO ENCOTRO NADA ***
            If FoundSomething = 0 Then
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .targetUser = 0
                .targetObj = 0
                .TargetObjMap = 0
                .TargetObjX = 0
                .TargetObjY = 0
                Call WriteMultiMessage(userIndex, eMessages.DontSeeAnything)
            End If
        End With
    Else
        If FoundSomething = 0 Then
            With .flags
                .TargetNPC = 0
                .TargetNpcTipo = eNPCType.Comun
                .targetUser = 0
                .targetObj = 0
                .TargetObjMap = 0
                .TargetObjX = 0
                .TargetObjY = 0
            End With
            
            Call WriteMultiMessage(userIndex, eMessages.DontSeeAnything)
        End If
    End If
End With

Exit Sub

ErrHandler:
    Call LogError("Error en LookAtTile. Error " & Err.Number & " : " & Err.description)

End Sub

Function FindDirection(Pos As WorldPos, Target As WorldPos) As eHeading
'***************************************************
'Author: Unknownn
'Last Modification: -
'Devuelve la direccion en la cual el target se encuentra
'desde pos, 0 si la direc es igual
'*****************************************************************

    Dim X As Integer
    Dim Y As Integer
    
    X = Pos.X - Target.X
    Y = Pos.Y - Target.Y
    
    'NE
    If Sgn(X) = -1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.NORTH, eHeading.EAST)
        Exit Function
    End If
    
    'NW
    If Sgn(X) = 1 And Sgn(Y) = 1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.NORTH)
        Exit Function
    End If
    
    'SW
    If Sgn(X) = 1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.WEST, eHeading.SOUTH)
        Exit Function
    End If
    
    'SE
    If Sgn(X) = -1 And Sgn(Y) = -1 Then
        FindDirection = IIf(RandomNumber(0, 1), eHeading.SOUTH, eHeading.EAST)
        Exit Function
    End If
    
    'Sur
    If Sgn(X) = 0 And Sgn(Y) = -1 Then
        FindDirection = eHeading.SOUTH
        Exit Function
    End If
    
    'norte
    If Sgn(X) = 0 And Sgn(Y) = 1 Then
        FindDirection = eHeading.NORTH
        Exit Function
    End If
    
    'oeste
    If Sgn(X) = 1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.WEST
        Exit Function
    End If
    
    'este
    If Sgn(X) = -1 And Sgn(Y) = 0 Then
        FindDirection = eHeading.EAST
        Exit Function
    End If
    
    'misma
    If Sgn(X) = 0 And Sgn(Y) = 0 Then
        FindDirection = 0
        Exit Function
    End If

End Function

Public Function ItemNoEsDeMapa(ByVal Index As Integer, ByVal bIsExit As Boolean) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: 27/12/2012 - ^[GS]^
'
'***************************************************

    With ObjData(Index)
        ItemNoEsDeMapa = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And Not (.OBJType = eOBJType.otTeleport And bIsExit) And Not (.OBJType = eOBJType.otDestruible)
    
    End With

End Function

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    With ObjData(Index)
        MostrarCantidad = .OBJType <> eOBJType.otPuertas And .OBJType <> eOBJType.otForos And .OBJType <> eOBJType.otCarteles And .OBJType <> eOBJType.otArboles And .OBJType <> eOBJType.otYacimiento And .OBJType <> eOBJType.otTeleport
    End With

End Function

Public Function EsObjetoFijo(ByVal OBJType As eOBJType) As Boolean
'***************************************************
'Author: Unknownn
'Last Modification: -
'
'***************************************************

    EsObjetoFijo = OBJType = eOBJType.otForos Or OBJType = eOBJType.otCarteles Or OBJType = eOBJType.otArboles Or OBJType = eOBJType.otYacimiento
End Function

Public Function RestrictStringToByte(ByRef restrict As String) As Byte ' 0.13.3
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    restrict = UCase$(restrict)
    
    Select Case restrict
        Case "NEWBIE"
            RestrictStringToByte = 1
            
        Case "ARMADA"
            RestrictStringToByte = 2
            
        Case "CAOS"
            RestrictStringToByte = 3
            
        Case "FACCION"
            RestrictStringToByte = 4
            
        Case Else
            RestrictStringToByte = 0
    End Select
End Function

Public Function RestrictByteToString(ByVal restrict As Byte) As String ' 0.13.3
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    Select Case restrict
        Case 1
            RestrictByteToString = "NEWBIE"
            
        Case 2
            RestrictByteToString = "ARMADA"
            
        Case 3
            RestrictByteToString = "CAOS"
            
        Case 4
            RestrictByteToString = "FACCION"
            
        Case 0
            RestrictByteToString = "NO"
    End Select
End Function

Public Function TerrainStringToByte(ByRef restrict As String) As Byte ' 0.13.3
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    restrict = UCase$(restrict)
    
    Select Case restrict
        Case "NIEVE"
            TerrainStringToByte = 1
            
        Case "DESIERTO"
            TerrainStringToByte = 2
            
        Case "CIUDAD"
            TerrainStringToByte = 3
            
        Case "CAMPO"
            TerrainStringToByte = 4
            
        Case "DUNGEON"
            TerrainStringToByte = 5
            
        Case Else
            TerrainStringToByte = 0
    End Select
End Function

Public Function TerrainByteToString(ByVal restrict As Byte) As String ' 0.13.3
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 04/18/2011
'
'***************************************************
    Select Case restrict
        Case 1
            TerrainByteToString = "NIEVE"
            
        Case 2
            TerrainByteToString = "DESIERTO"
            
        Case 3
            TerrainByteToString = "CIUDAD"
            
        Case 4
            TerrainByteToString = "CAMPO"
            
        Case 5
            TerrainByteToString = "DUNGEON"
            
        Case 0
            TerrainByteToString = "BOSQUE"
    End Select
End Function

Public Function SEncriptar(ByVal Cadena As String) As String
' GSZ-AO - Encripta una cadena de texto
    Dim i As Long, RandomNum As Integer
    
    RandomNum = 99 * Rnd
    If RandomNum < 10 Then RandomNum = 10
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
    Next i
    SEncriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
    DoEvents

End Function

Public Function SDesencriptar(ByVal Cadena As String) As String
' GSZ-AO - Desencripta una cadena de texto
    Dim i As Long, NumDesencriptar As String
    
    NumDesencriptar = Chr$(Asc(Left$((Right(Cadena, 2)), 1)) - 10) & Chr$(Asc(Right$((Right(Cadena, 2)), 1)) - 10)
    Cadena = (Left$(Cadena, Len(Cadena) - 2))
    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) - NumDesencriptar)
    Next i
    SDesencriptar = Cadena
    DoEvents

End Function

' GSZAO - Encriptaci?n basica y rapida para Strings
Public Function RndCrypt(ByVal str As String, ByVal Password As String) As String
    '  Made by Michael Ciurescu
    ' (CVMichael from vbforums.com)
    '  Original thread: http://www.vbforums.com/showthread.php?t=231798
    Dim SK As Long, k As Long

    Rnd -1
    Randomize Len(Password)

    For k = 1 To Len(Password)
        SK = SK + (((k Mod 256) _
        Xor Asc(mid$(Password, k, 1))) _
        Xor Fix(256 * Rnd))
    Next k

    Rnd -1
    Randomize SK
    
    For k = 1 To Len(str)
        Mid$(str, k, 1) = Chr(Fix(256 * Rnd) _
        Xor Asc(mid$(str, k, 1)))
    Next k
    
    RndCrypt = str
End Function


Public Function QuitarTildes(ByVal sString As String) As String
 
    QuitarTildes = Replace$(Replace$(Replace$(Replace$(Replace$(UCase$(sString), "?", "A"), "?", "E"), "?", "I"), "?", "O"), "?", "U")
 
End Function

