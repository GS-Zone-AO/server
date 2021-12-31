Attribute VB_Name = "modUsuariosComercioSeguro"
'**************************************************************
' mdlComercioConUsuarios.bas - Allows players to commerce between themselves.
'
' Designed and implemented by Alejandro Santos (AlejoLP)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

'[Alejo]
Option Explicit

Public Const MAX_OFFER_SLOTS As Integer = 20
Public Const GOLD_OFFER_SLOT As Integer = MAX_OFFER_SLOTS + 1

Public Type tCOmercioUsuario
    DestUsu As Integer 'El otro Usuario
    DestNick As String
    Objeto(1 To MAX_OFFER_SLOTS) As Integer 'Indice de los objetos que se desea dar
    GoldAmount As Long
    
    cant(1 To MAX_OFFER_SLOTS) As Long 'Cuantos objetos desea dar
    Acepto As Boolean
    Confirmo As Boolean
End Type

Private Type tOfferItem ' 0.13.3
    ObjIndex As Integer
    Amount As Long
End Type

'origen: origen de la transaccion, originador del comando
'destino: receptor de la transaccion
Public Sub IniciarComercioConUsuario(ByVal Origen As Integer, ByVal Destino As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 10/07/2012 - ^[GS]^
'
'***************************************************
    On Error GoTo ErrHandler
    
    'Si ambos pusieron /comerciar entonces
    If UserList(Origen).ComUsu.DestUsu = Destino And UserList(Destino).ComUsu.DestUsu = Origen Then
        'Actualiza el inventario del usuario
        Call UpdateUserInv(True, Origen, 0)
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Origen)
        UserList(Origen).flags.Comerciando = True
    
        If UserList(Origen).flags.Comerciando Or UserList(Destino).flags.Comerciando Then
            Call WriteMensajes(Origen, eMensajes.Mensaje452) ' "No puedes comerciar en este momento"
            Call WriteMensajes(Destino, eMensajes.Mensaje452)
            Exit Sub
        End If
    
        'Actualiza el inventario del usuario
        Call UpdateUserInv(True, Destino, 0)
        'Decirle al origen que abra la ventanita.
        Call WriteUserCommerceInit(Destino)
        UserList(Destino).flags.Comerciando = True
    
        'Call EnviarObjetoTransaccion(Origen)
    Else
        'Es el primero que comercia ?
        Call WriteConsoleMsg(Destino, UserList(Origen).Name & " desea comerciar. Si deseas aceptar, escribe /COMERCIAR.", FontTypeNames.FONTTYPE_TALK)
        UserList(Destino).flags.targetUser = Origen
        
    End If
    
    Call FlushBuffer(Destino)
    
    Exit Sub
ErrHandler:
        Call LogError("Error en IniciarComercioConUsuario: " & Err.description)
End Sub

Public Sub EnviarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte)
'***************************************************
'Autor: Unknown
'Last Modification: 10/07/2012 - ^[GS]^
'Sends the offer change to the other trading user
'***************************************************
On Error GoTo ErrHandler

    Dim ObjIndex As Integer
    Dim ObjAmount As Long
    Dim OtherUserIndex As Integer
    
    OtherUserIndex = UserList(UserIndex).ComUsu.DestUsu
    
    With UserList(OtherUserIndex)
        If OfferSlot = GOLD_OFFER_SLOT Then
            ObjIndex = iORO
            ObjAmount = .ComUsu.GoldAmount
        Else
            ObjIndex = .ComUsu.Objeto(OfferSlot)
            ObjAmount = .ComUsu.cant(OfferSlot)
        End If
    End With
    
    Exit Sub

ErrHandler:
    LogError "Error en EnviarOferta. Error: " & Err.description & ". UserIndex: " & UserIndex & ". OfferSlot: " & OfferSlot
  
End Sub

Public Sub FinComerciarUsu(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************
On Error GoTo ErrHandler
    
    Dim i As Long
    
    With UserList(UserIndex)
        If .ComUsu.DestUsu > 0 And .ComUsu.DestUsu < iniMaxUsuarios Then ' 0.13.5
            Call WriteUserCommerceEnd(UserIndex)
        End If
        
        .ComUsu.Acepto = False
        .ComUsu.Confirmo = False
        .ComUsu.DestUsu = 0
        
        For i = 1 To MAX_OFFER_SLOTS
            .ComUsu.cant(i) = 0
            .ComUsu.Objeto(i) = 0
        Next i
        
        .ComUsu.GoldAmount = 0
        .ComUsu.DestNick = vbNullString
        .flags.Comerciando = False
    End With
    
    Exit Sub
    
ErrHandler:
    LogError "Error en FinComerciarUsu. Error: " & Err.description & ". UserIndex: " & UserIndex

End Sub

Public Sub AceptarComercioUsu(ByVal UserIndex As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 27/07/2012 - ^[GS]^
'***************************************************
    Dim TradingObj As Obj
    Dim OtroUserIndex As Integer
    Dim TerminarAhora As Boolean
    Dim OfferSlot As Integer

    UserList(UserIndex).ComUsu.Acepto = True
    
    OtroUserIndex = UserList(UserIndex).ComUsu.DestUsu
    
    ' User valido?
    If OtroUserIndex <= 0 Or OtroUserIndex > iniMaxUsuarios Then ' 0.13.3
        Call FinComerciarUsu(UserIndex)
        Exit Sub
    End If
    
    ' Acepto el otro?
    If UserList(OtroUserIndex).ComUsu.Acepto = False Then
        Exit Sub
    End If
    
    ' Aceptaron ambos, chequeo que tengan los items que ofertaron
    If Not HasOfferedItems(UserIndex) Then ' 0.13.3
        
        Call WriteConsoleMsg(UserIndex, "���El comercio se cancel� porque no posees los �tems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(OtroUserIndex, "���El comercio se cancel� porque " & UserList(UserIndex).Name & " no posee los �tems que ofert�!!!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FinComerciarUsu(UserIndex)
        Call FinComerciarUsu(OtroUserIndex)
        Call modProtocol.FlushBuffer(OtroUserIndex)
        
        Exit Sub
        
    ElseIf Not HasOfferedItems(OtroUserIndex) Then ' 0.13.3
        
        Call WriteConsoleMsg(UserIndex, "���El comercio se cancel� porque " & UserList(OtroUserIndex).Name & " no posee los �tems que ofert�!!!", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(OtroUserIndex, "���El comercio se cancel� porque no posees los �tems que ofertaste!!!", FontTypeNames.FONTTYPE_FIGHT)
        
        Call FinComerciarUsu(UserIndex)
        Call FinComerciarUsu(OtroUserIndex)
        Call modProtocol.FlushBuffer(OtroUserIndex)
        
        Exit Sub
        
    End If
    
    
    ' Envio los items a quien corresponde
    For OfferSlot = 1 To MAX_OFFER_SLOTS + 1
        
        ' Items del 1er usuario
        With UserList(UserIndex)
            ' Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                ' Quito la cantidad de oro ofrecida
                .Stats.GLD = .Stats.GLD - .ComUsu.GoldAmount
                ' Log
                If .ComUsu.GoldAmount >= MIN_GOLD_AMOUNT_LOG Then Call LogDesarrollo(.Name & " solt� oro en comercio seguro con " & UserList(OtroUserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Usuario
                Call WriteUpdateGold(UserIndex) ' 0.13.5
                ' Se la doy al otro
                UserList(OtroUserIndex).Stats.GLD = UserList(OtroUserIndex).Stats.GLD + .ComUsu.GoldAmount
                ' Update Otro Usuario
                Call WriteUpdateGold(OtroUserIndex) ' 0.13.5
                
            ' Le pasa lo ofertado de los slots con items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                
                'Quita el objeto y se lo da al otro
                If Not MeterItemEnInventario(OtroUserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(OtroUserIndex).Pos, TradingObj)
                End If
            
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, UserIndex)
                
                'Es un Objeto que tenemos que loguear?
                If ((ObjData(TradingObj.ObjIndex).Log = 1) Or (ObjData(TradingObj.ObjIndex).OBJType = eOBJType.otLlaves)) Then ' 0.13.5
                    Call LogDesarrollo(.Name & " le pas� en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                'Es mucha cantidad?
                ElseIf TradingObj.Amount >= MIN_AMOUNT_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pas� en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                ElseIf (TradingObj.Amount * ObjData(TradingObj.ObjIndex).Valor) >= MIN_VALUE_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pas� en comercio seguro a " & UserList(OtroUserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                End If
            End If
        End With
        
        ' Items del 2do usuario
        With UserList(OtroUserIndex)
            ' Le pasa el oro
            If OfferSlot = GOLD_OFFER_SLOT Then
                ' Quito la cantidad de oro ofrecida
                .Stats.GLD = .Stats.GLD - .ComUsu.GoldAmount
                ' Log
                If .ComUsu.GoldAmount >= MIN_GOLD_AMOUNT_LOG Then Call LogDesarrollo(.Name & " solt� oro en comercio seguro con " & UserList(UserIndex).Name & ". Cantidad: " & .ComUsu.GoldAmount)
                ' Update Usuario
                Call WriteUpdateGold(OtroUserIndex) ' 0.13.5
                'y se la doy al otro
                UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + .ComUsu.GoldAmount
                ' Update Otro Usuario
                Call WriteUpdateGold(UserIndex) ' 0.13.5
                
            ' Le pasa la oferta de los slots con items
            ElseIf .ComUsu.Objeto(OfferSlot) > 0 Then
                TradingObj.ObjIndex = .ComUsu.Objeto(OfferSlot)
                TradingObj.Amount = .ComUsu.cant(OfferSlot)
                
                'Quita el objeto y se lo da al otro
                If Not MeterItemEnInventario(UserIndex, TradingObj) Then
                    Call TirarItemAlPiso(UserList(UserIndex).Pos, TradingObj)
                End If
            
                Call QuitarObjetos(TradingObj.ObjIndex, TradingObj.Amount, OtroUserIndex)
                
                'Es un Objeto que tenemos que loguear?
                If ((ObjData(TradingObj.ObjIndex).Log = 1) Or (ObjData(TradingObj.ObjIndex).OBJType = eOBJType.otLlaves)) Then ' 0.13.5
                    Call LogDesarrollo(.Name & " le pas� en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                'Es mucha cantidad?
                ElseIf TradingObj.Amount >= MIN_AMOUNT_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pas� en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                ElseIf (TradingObj.Amount * ObjData(TradingObj.ObjIndex).Valor) >= MIN_VALUE_LOG Then
                    'Si no es de los prohibidos de loguear, lo logueamos.
                    If ObjData(TradingObj.ObjIndex).NoLog <> 1 Then
                        Call LogDesarrollo(.Name & " le pas� en comercio seguro a " & UserList(UserIndex).Name & " " & TradingObj.Amount & " " & ObjData(TradingObj.ObjIndex).Name)
                    End If
                End If
            End If
        End With
        
    Next OfferSlot

    ' End Trade
    Call FinComerciarUsu(UserIndex)
    Call FinComerciarUsu(OtroUserIndex)
    Call modProtocol.FlushBuffer(OtroUserIndex) ' 0.13.3

End Sub

Public Sub AgregarOferta(ByVal UserIndex As Integer, ByVal OfferSlot As Byte, ByVal ObjIndex As Integer, ByVal Amount As Long, ByVal IsGold As Boolean)
'***************************************************
'Autor: ZaMa
'Last Modification: 27/07/2012 - ^[GS]^
'Adds gold or items to the user's offer
'***************************************************
On Error GoTo ErrHandler

    With UserList(UserIndex).ComUsu ' 0.13.5
        ' Si ya confirmo su oferta, no puede cambiarla!
        If Not .Confirmo Then
            If IsGold Then
            ' Agregamos (o quitamos) mas oro a la oferta
                .GoldAmount = .GoldAmount + Amount
                
                ' Imposible que pase, pero por las dudas..
                If .GoldAmount < 0 Then .GoldAmount = 0
            Else
            ' Agreamos (o quitamos) el item y su cantidad en el slot correspondiente
                ' Si es 0 estoy modificando la cantidad, no agregando
                If ObjIndex > 0 Then .Objeto(OfferSlot) = ObjIndex
                .cant(OfferSlot) = .cant(OfferSlot) + Amount
                
                'Quit� todos los items de ese tipo
                If .cant(OfferSlot) <= 0 Then
                    ' Removemos el objeto para evitar conflictos
                    .Objeto(OfferSlot) = 0
                    .cant(OfferSlot) = 0
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrHandler:

    LogError "Error en AgregarOferta. Error: " & Err.description & ". UserIndex: " & UserIndex
    
End Sub

Public Function PuedeSeguirComerciando(ByVal UserIndex As Integer) As Boolean
'***************************************************
'Autor: ZaMa
'Last Modification: 27/07/2012 - ^[GS]^
'Validates wether the conditions for the commerce to keep going are satisfied
'***************************************************
On Error GoTo ErrHandler

    Dim OtroUserIndex As Integer
    Dim ComercioInvalido As Boolean

    With UserList(UserIndex) ' 0.13.5
    
        OtroUserIndex = .ComUsu.DestUsu
        
        ' Usuario valido?
        If OtroUserIndex <= 0 Or OtroUserIndex > iniMaxUsuarios Then
            ComercioInvalido = True
        End If
        
        If Not ComercioInvalido Then
            ' Estan logueados?
            If UserList(OtroUserIndex).flags.UserLogged = False Or .flags.UserLogged = False Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' Se estan comerciando el uno al otro?
            If UserList(OtroUserIndex).ComUsu.DestUsu <> UserIndex Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' El nombre del otro es el mismo que al que le comercio?
            If UserList(OtroUserIndex).Name <> .ComUsu.DestNick Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' Mi nombre  es el mismo que al que el le comercia?
            If .Name <> UserList(OtroUserIndex).ComUsu.DestNick Then
                ComercioInvalido = True
            End If
        End If
        
        If Not ComercioInvalido Then
            ' Esta vivo?
            If UserList(OtroUserIndex).flags.Muerto = 1 Then
                ComercioInvalido = True
            End If
        End If
        
        ' Fin del comercio
        If ComercioInvalido = True Then
            Call FinComerciarUsu(UserIndex)
            
            If OtroUserIndex > 0 And OtroUserIndex <= iniMaxUsuarios Then
                Call FinComerciarUsu(OtroUserIndex)
                Call FlushBuffer(OtroUserIndex)
            End If
            
            Exit Function
        End If
    End With
    
    PuedeSeguirComerciando = True
    
    Exit Function

ErrHandler:

    LogError "Error en PuedeSeguirComerciando. Error: " & Err.description & ". UserIndex: " & UserIndex

End Function

Private Function HasOfferedItems(ByVal UserIndex As Integer) As Boolean ' 0.13.3
'***************************************************
'Autor: ZaMa
'Last Modification: 10/07/2012 - ^[GS]^
'Checks whether the user has the offered items in his inventory or not.
'***************************************************

    Dim OfferedItems(MAX_OFFER_SLOTS - 1) As tOfferItem
    Dim Slot As Long
    Dim SlotAux As Long
    Dim SlotCount As Long
    
    Dim ObjIndex As Integer
    
    With UserList(UserIndex).ComUsu
        
        ' Agrupo los items que son iguales
        For Slot = 1 To MAX_OFFER_SLOTS
                    
            ObjIndex = .Objeto(Slot)
            
            If ObjIndex > 0 Then
            
                For SlotAux = 0 To SlotCount - 1
                    
                    If ObjIndex = OfferedItems(SlotAux).ObjIndex Then
                        ' Son iguales, aumento la cantidad
                        OfferedItems(SlotAux).Amount = OfferedItems(SlotAux).Amount + .cant(Slot)
                        Exit For
                    End If
                    
                Next SlotAux
                
                ' No encontro otro igual, lo agrego
                If SlotAux = SlotCount Then
                    OfferedItems(SlotCount).ObjIndex = ObjIndex
                    OfferedItems(SlotCount).Amount = .cant(Slot)
                    
                    SlotCount = SlotCount + 1
                End If
                
            End If
            
        Next Slot
        
        ' Chequeo que tengan la cantidad en el inventario
        For Slot = 0 To SlotCount - 1
            If Not HasEnoughItems(UserIndex, OfferedItems(Slot).ObjIndex, OfferedItems(Slot).Amount) Then Exit Function
        Next Slot
        
        ' Compruebo que tenga el oro que oferta
        If UserList(UserIndex).Stats.GLD < .GoldAmount Then Exit Function
        
    End With
    
    HasOfferedItems = True

End Function
