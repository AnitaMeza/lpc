Imports System.Data.SqlClient
Imports System.Data
Partial Class Login
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'txtContrasenia.Attributes.Add("onfocus", "this.type='password';")
    End Sub


    ''' <summary>
    ''' OBSOLETO PARA LA NUEVA VERSIÓN, ESTE ÚNICAMENTE SERVIRÁ PARA SABER SI LA PERSONA QUE SE LOGUEA ESTÁ EN EL SISTEMA ANTERIOR Y REDIRECCIONARLO
    ''' </summary>
    ''' <param name="IdAdmin"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ObtenerMembresia(ByVal IdAdmin As Integer) As String
        Dim CN As New czCN
        Try
            Dim sen As New czSentencias
            Dim cmd As New SqlCommand
            cmd.CommandText = "Select top 1 * from PruebasCompetencias Where UsuarioId = @IdAdmin"
            cmd.Parameters.AddWithValue("@IdAdmin", IdAdmin)
            Dim dt As Data.DataTable = sen.ObtenerDataTable(cmd, czCon.cnPLC)

            Dim TieneLPC As Integer = 0
            Dim TienePVC As Integer = 0
            Dim TienePERS As Integer = 0

            ''LA SESSION LLAMADA Pruebas, tiene los siguientes valores:
            ''Valor de inicialización (no tiene pruebas)=0
            'LPC(1,0,0)=1
            'PVC(0,2,0)=2
            'LPC Y PVC(1,2,0)=3   
            'personalizado(0,0,4)=4
            'lpc y Personalizado(1,0,4)=5
            'pvc y personalizado(0,2,4)=6
            'todos(1,2,4)=7          

            With dt.Rows(0)
                If CBool(.Item("LPC").ToString()) = True Then
                    TieneLPC = 1
                End If
                If CBool(.Item("PVC").ToString()) = True Then
                    TienePVC = 2
                End If
                If (CBool(.Item("Personalizado").ToString)) = True Then
                    TienePERS = 4
                End If

                Session.Add("Pruebas", CInt(TieneLPC + TienePVC + TienePERS))

                Dim fecVen As New DateTime
                Dim varPruebas As Integer = Integer.Parse(Session.Item("Pruebas").ToString)

                Select Case varPruebas ''Session.Item("Pruebas")
                    Case 1  'LPC
                        fecVen = New DateTime(Year(.Item("FechaFinLPC")), Month(.Item("FechaFinLPC")), Day(.Item("FechaFinLPC")))
                        If (DateTime.Now > fecVen) Then
                            Session.Add("Pruebas", 0)
                            'YA EXPIRÓ LA MEMBRESÍA
                            Return "Lo sentimos, <br/> su membresía de LPC ha caducado"
                        End If
                    Case 2  'PVC
                        fecVen = New DateTime(Year(.Item("FechaFinPVC")), Month(.Item("FechaFinPVC")), Day(.Item("FechaFinPVC")))
                        If (DateTime.Now > fecVen) Then
                            Session.Add("Pruebas", 0)
                            'YA EXPIRÓ LA MEMBRESÍA
                            Return "Lo sentimos, <br/> su membresía de PVC ha caducado"
                        End If
                    Case 3  'LPC Y PVC
                        fecVen = New DateTime(Year(.Item("FechaFinLPC")), Month(.Item("FechaFinLPC")), Day(.Item("FechaFinLPC")))
                        Dim fecVenPVC As DateTime = New DateTime(Year(.Item("FechaFinPVC")), Month(.Item("FechaFinPVC")), Day(.Item("FechaFinPVC")))
                        If (DateTime.Now > fecVen) Then '¿ VENCIÓ LPC ?
                            Session.Add("Pruebas", 2)   'SÓLO TENDRÁ PVC
                        End If
                        If (DateTime.Now > fecVenPVC) Then '¿ VENCIÓ PVC ?
                            If (Session.Item("Pruebas") = 2) Then ' SI SOLO TIENE LA PVC, ENTONCES EN ESTE MOMENTO LE QUITAREMOS PVC Y SE QUEDARÁ SIN ADMISIÓN
                                Session.Add("Pruebas", 0)   'SI LLEGA AQUÍ NO TENDRÁ NI LPC NI PVC
                                'YA EXPIRÓ LA MEMBRESÍA
                                Return "Lo sentimos, <br/> sus membresías han caducado"
                            Else
                                Session.Add("Pruebas", 1)   'SI LLEGA AQUÍ SOLO TENDRÁ LPC
                            End If
                        End If
                    Case 4 'personalizado
                        fecVen = New DateTime(Year(.Item("FechaFinPers")), Month(.Item("FechaFinPers")), Day(.Item("FechaFinPers")))
                        If (DateTime.Now > fecVen) Then
                            Session.Add("Pruebas", 0)
                            'YA EXPIRÓ LA MEMBRESÍA
                            Return "Lo sentimos, <br/> su membresía Personalizada ha caducado"
                        End If
                    Case 5  'LPC Y Personalizado
                        fecVen = New DateTime(Year(.Item("FechaFinLPC")), Month(.Item("FechaFinLPC")), Day(.Item("FechaFinLPC")))
                        Dim fecVenPers As DateTime = New DateTime(Year(.Item("FechaFinPers")), Month(.Item("FechaFinPers")), Day(.Item("FechaFinPers")))
                        If (DateTime.Now > fecVen) Then '¿ VENCIÓ LPC ?
                            Session.Add("Pruebas", 4)   'SÓLO TENDRÁ PERSONALIZADO
                        End If
                        If (DateTime.Now > fecVenPers) Then '¿ VENCIÓ PERSONALIZADO ?
                            If (Session.Item("Pruebas") = 4) Then ' SI SOLO TIENE PERSONALIZADO, ENTONCES EN ESTE MOMENTO LE QUITAREMOS PVC Y SE QUEDARÁ SIN ADMISIÓN
                                Session.Add("Pruebas", 0)   'SI LLEGA AQUÍ NO TENDRÁ NI LPC NI PERSONALIZADO
                                'YA EXPIRÓ LA MEMBRESÍA
                                Return "Lo sentimos, <br/> sus membresías han caducado"
                            Else
                                Session.Add("Pruebas", 1)   'SI LLEGA AQUÍ SOLO TENDRÁ LPC
                            End If
                        End If
                    Case 6  'PVC y perzonalido
                        fecVen = New DateTime(Year(.Item("FechaFinPers")), Month(.Item("FechaFinPers")), Day(.Item("FechaFinPers")))
                        Dim fecVenPVC As DateTime = New DateTime(Year(.Item("FechaFinPVC")), Month(.Item("FechaFinPVC")), Day(.Item("FechaFinPVC")))
                        If (DateTime.Now > fecVen) Then '¿ VENCIÓ PERSONALIZADO ?
                            Session.Add("Pruebas", 2)   'SÓLO TENDRÁ PVC
                        End If
                        If (DateTime.Now > fecVenPVC) Then '¿ VENCIÓ PVC ?
                            If (Session.Item("Pruebas") = 2) Then ' SI SOLO TIENE LA PVC, ENTONCES EN ESTE MOMENTO LE QUITAREMOS PVC Y SE QUEDARÁ SIN ADMISIÓN
                                Session.Add("Pruebas", 0)   'SI LLEGA AQUÍ NO TENDRÁ NI LPC NI PVC
                                'YA EXPIRÓ LA MEMBRESÍA
                                Return "Lo sentimos, <br/> sus membresías han caducado"
                            Else
                                Session.Add("Pruebas", 4)   'SI LLEGA AQUÍ SOLO TENDRÁ LPC
                            End If
                        End If
                    Case 7  'LPC, PVC y personalizado 
                        fecVen = New DateTime(Year(.Item("FechaFinLPC")), Month(.Item("FechaFinLPC")), Day(.Item("FechaFinLPC")))
                        Dim fecVenPVC As DateTime = New DateTime(Year(.Item("FechaFinPVC")), Month(.Item("FechaFinPVC")), Day(.Item("FechaFinPVC")))
                        Dim fecVenPers As DateTime = New DateTime(Year(.Item("FechaFinPers")), Month(.Item("FechaFinPers")), Day(.Item("FechaFinPers")))
                        If (DateTime.Now > fecVen) Then '¿ VENCIÓ LPC ?
                            Session.Add("Pruebas", 6)   'SÓLO TENDRÁ PVC y PERSONALIZADO
                        End If
                        If (DateTime.Now > fecVenPVC) Then '¿ VENCIÓ PVC ?
                            If (Session.Item("Pruebas") = 6) Then ' SI SOLO TIENE LA PVC, ENTONCES EN ESTE MOMENTO LE QUITAREMOS PVC Y SE QUEDARÁ SIN ADMISIÓN
                                Session.Add("Pruebas", 4)   'SI LLEGA AQUÍ NO TENDRÁ NI LPC NI PVC
                            Else
                                Session.Add("Pruebas", 5)   'SI LLEGA AQUÍ SOLO TENDRÁ LPC y PERSONALIZADO
                            End If
                        End If
                        If (DateTime.Now > fecVenPers) Then '¿ VENCIÓ PERSONALIZADO?
                            If (Session.Item("Pruebas") = 7) Then
                                Session.Add("Pruebas", 3) 'SI LLEGA AQUÍ SOLO TENDRÁ LPC y PVC
                            ElseIf (Session.Item("Pruebas") = 6) Then
                                Session.Add("Pruebas", 2) 'SI LLEGA AQUÍ SOLO TENDRÁ PVC
                            ElseIf (Session.Item("Pruebas") = 5) Then
                                Session.Add("Pruebas", 1) 'SI LLEGA AQUÍ SOLO TENDRÁ LPC
                            Else
                                Session.Add("Pruebas", 0)
                                'YA EXPIRÓ LA MEMBRESÍA
                                Return "Lo sentimos, <br/> su membresía Personalizada ha caducado"
                            End If
                        End If
                    Case Else
                        Return "Lo sentimos, usted no tiene membresía"
                End Select
            End With
        Catch ex As Exception
            Return ex.Message
        End Try
        Return String.Empty
    End Function

    Public Nivel As Integer
    Public IdUsuario As Integer

    Private Function valida(ByVal usuario As String, ByVal pass As String) As String
        Me.LblMensajes.Visible = False
        Dim mensaje As String = String.Empty
        Dim dt As New DataTable
        Dim CN As New czCN
        'La variable de sesión que nos indica que el usuario esta logueado es UsuarioId
        'La variable de sesión que nos indica de que nivel es: NivelId
        'Hay 3 niveles, 1=administrador, 2=administrativo, 3=aplicador

        'Verificamos si es administrador
        Dim daAdmin As New SqlDataAdapter("SELECT TOP 1 ID, LPC, LPCFF,Competencias, Entrevista, Paquete, IdPaquete, TotalCompetenciasUsuario, Espanol, Ingles, Portugues FROM lpc_vsAdministradores where Usuario=@usuario and Contrasenia=@pass", czCN.obtenerCadenaPsicoWeb())
        daAdmin.SelectCommand.Parameters.AddWithValue("@usuario", usuario)
        daAdmin.SelectCommand.Parameters.AddWithValue("@pass", pass)
        daAdmin.Fill(dt)
        If (dt.Rows.Count > 0) Then
            'SON USUARIOS ADMINISTRADORES
            If (CBool(dt.Rows(0).Item("LPC")) = True) Then
                If (CDate(dt.Rows(0).Item("LPCFF")) > DateTime.Now) Then ' si es mayor la fecha de vencimiento, aun puede usarlo
                    Session.Add("NivelId", 1)
                    Session.Add("UsuarioId", dt.Rows(0).Item("Id"))
                    Dim version As New czVersionLPC
                    version.Competencias = dt.Rows(0)("Competencias").ToString()
                    version.Entrevista = CBool(dt.Rows(0)("Entrevista"))
                    version.Predeterminadas = True
                    version.Paquete = dt.Rows(0)("IdPaquete").ToString()
                    version.Version = dt.Rows(0)("Paquete").ToString()
                    version.Español = CBool(dt.Rows(0)("Espanol").ToString())
                    version.Portugues = CBool(dt.Rows(0)("Portugues").ToString())
                    version.Ingles = CBool(dt.Rows(0)("Ingles").ToString())
                    Session.Add("Version", version)
                    Return String.Empty
                Else
                    mensaje = "La membresía ha terminado para este usuario, póngase en contacto con su proveedor"
                End If
            Else
                mensaje = "El usuario no tiene membresía"
            End If
        Else
            'VERIFICAR SI SON USUARIOS DE APLICACIÓN
            daAdmin = New SqlDataAdapter("SELECT TOP 1 * FROM lpc_vsAplicadores where Usuario=@usuario and Contrasenia=@pass", czCN.obtenerCadenaPsicoWeb())
            daAdmin.SelectCommand.Parameters.AddWithValue("@usuario", usuario)
            daAdmin.SelectCommand.Parameters.AddWithValue("@pass", pass)
            daAdmin.Fill(dt)
            If (dt.Rows.Count > 0) Then
                If (Now > CDate(dt.Rows(0).Item("LPCFF").ToString())) Then ' VERIFICAMOS LA FECHA DE VENCIMIENTO DEL USUARIO GENERAL
                    'mensaje = "La membresía ha caducado"
                    Return "La membresía ha caducado"
                Else
                    If (CBool(dt.Rows(0).Item("Caduca").ToString())) Then
                        If (Now > CDate(dt.Rows(0).Item("FechaFinal").ToString())) Then ' NO PUEDE ACCESAR
                            'mensaje = "La contraseña ha caducado"
                            Return "La contraseña ha caducado"
                        End If
                    Else
                        Dim sen As New czSentencias
                        Session.Add("Expediente", dt.Rows(0)("Cedula"))
                        Session.Add("ConExp", dt.Rows(0)("ConExp"))
                        Session.Add("UsuarioId", dt.Rows(0)("IdUsuario"))
                        Session.Add("UserId", dt.Rows(0)("UserId"))

                        Dim cmd As New Data.SqlClient.SqlCommand
                        cmd.CommandText = "Select EnviarCorreoLPC, todosLPC from CorreoN4 where idUsuario = @Id"
                        cmd.Parameters.AddWithValue("@Id", Session.Item("UsuarioId"))
                        Dim dt0 As Data.DataTable = sen.ObtenerDataTable(cmd, czCN.obtenerCadenaPsicoWeb())
                        If (dt0.Rows.Count > 0) Then
                            If CBool(dt0.Rows(0)("EnviarCorreoLPC")) Then
                                If CBool(dt0.Rows(0)("todosLPC")) Then
                                    Session.Add("EnvioCorreo", "True")
                                Else
                                    Session.Add("EnvioCorreo", dt.Rows(0)("EnvioCorreo"))
                                End If
                            Else
                                Session.Add("EnvioCorreo", "False")
                            End If
                        End If

                        Session.Add("EnvioCorreo", dt.Rows(0)("EnvioCorreo"))
                        Session.Add("NivelId", IIf(CInt(dt.Rows(0)("tipo2")) > 2, 4, CInt(dt.Rows(0)("tipo2"))))
                        Session.Add("Puesto", dt.Rows(0)("Puesto"))
                        If (dt.Rows(0)("Nivel1") = True) Then
                            Session.Add("Examen", 1)
                        End If
                        If (dt.Rows(0)("Nivel2") = True) Then
                            Session.Add("Examen", 2)
                        End If
                        If (dt.Rows(0)("Nivel3") = True) Then
                            Session.Add("Examen", 3)
                        End If
                        If (dt.Rows(0)("Nivel4") = True) Then
                            Session.Add("Examen", 4)
                        End If
                        If (dt.Rows(0)("Nivel5") = True) Then
                            Session.Add("Examen", 5)
                        End If

                        Dim dt1 As Data.DataTable = sen.ObtenerDataTable("SELECT IdPaquete, Paquete, TotalCompetencias, SeleccionCompetencias, Entrevista, IdLPC, IdUsuario, Competencias, Actualizacion, TotalCompetenciasUsuario FROM lpc_vsVersionLPC Where IdUsuario = " & Session.Item("UsuarioId"), czCN.obtenerCadena(czCon.Pw))
                        Dim version As New czVersionLPC
                        version.Competencias = dt1.Rows(0)("Competencias").ToString()
                        version.Entrevista = CBool(dt1.Rows(0)("Entrevista"))
                        version.Predeterminadas = True
                        version.Paquete = dt1.Rows(0)("IdPaquete").ToString()
                        version.Version = dt1.Rows(0)("Paquete").ToString()
                        Session.Add("Version", version)

                        Return String.Empty
                        'Return ObtenerMembresia(dt.Rows(0)("Id"))
                    End If
                End If
            Else
                'VERIFICAMOS SI ESTE USUARIO ES DEL SISTEMA PASADO

                '====================================================================================
                'Verificamos si es administrador
                daAdmin = New SqlDataAdapter("select IdAdmin from Admin where Usuario=@usuario and Pass=@pass", czCon.cnPLC)
                daAdmin.SelectCommand.Parameters.AddWithValue("@usuario", usuario)
                daAdmin.SelectCommand.Parameters.AddWithValue("@pass", pass)
                Dim DS As New DataSet
                Try
                    daAdmin.Fill(DS, "Admin")
                    If DS.Tables("Admin").Rows.Count <> 0 Then
                        Me.Session.Add("UsuarioId", DS.Tables("Admin").Rows(0)("IdAdmin"))
                        Me.Session.Add("NivelId", 1)
                        If (ObtenerMembresia(DS.Tables("Admin").Rows(0)("IdAdmin")) = String.Empty) Then
                            Session.Clear()
                            Dim seg As New czSeguridadQuery()
                            Dim u As String = seg.Cifrar(Me.txtUsuario.Text)
                            Dim p As String = seg.Cifrar(Me.txtContrasenia.Text)
                            Response.Redirect(czCon.LinkR & "?u=" & u & "&p=" & p)
                        End If
                    End If

                    'Verificamos si es de nivel 2 o 3
                    Dim daUsr As New SqlDataAdapter("select IdUsuario, IdAdmin, Nivel1, Nivel2, Nivel3, Nivel4, ConExp, Expediente, Tipo2,Puesto from usuarios where Usuario=@usuario and Pass=@pass", czCN.obtenerCadena(czCon.LPC))
                    daUsr.SelectCommand.Parameters.AddWithValue("@usuario", usuario)
                    daUsr.SelectCommand.Parameters.AddWithValue("@pass", pass)
                    daUsr.Fill(DS, "Usr")

                    If DS.Tables("Usr").Rows.Count <> 0 Then
                        Session.Add("Expediente", DS.Tables("Usr").Rows(0)("Expediente"))
                        Session.Add("ConExp", DS.Tables("Usr").Rows(0)("ConExp"))
                        Session.Add("UsuarioId", DS.Tables("Usr").Rows(0)("IdAdmin"))
                        Session.Add("UserId", DS.Tables("Usr").Rows(0)("IdUsuario"))
                        Session.Add("NivelId", DS.Tables("Usr").Rows(0)("tipo2"))
                        Session.Add("Puesto", DS.Tables("Usr").Rows(0)("Puesto"))
                        If (ObtenerMembresia(DS.Tables("Usr").Rows(0)("IdAdmin")) = String.Empty) Then
                            Session.Clear()
                            Dim seg As New czSeguridadQuery()
                            Dim u As String = seg.Cifrar(Me.txtUsuario.Text)
                            Dim p As String = seg.Cifrar(Me.txtContrasenia.Text)
                            Response.Redirect(czCon.LinkR & "?u=" & u & "&p=" & p)
                        End If
                    Else
                        'mensaje = "El usuario NO existe"
                        Return "El usuario NO existe"
                    End If
                Catch ex As Exception
                    'mensaje = "El usuario NO existe"
                    Return "El usuario NO existe"
                End Try
                '====================================================================================
                'mensaje = "El usuario NO existe"
            End If
        End If
        Return "Los datos son incorrectos, <br/> por favor vuelva a intentarlo"
    End Function

    Protected Sub btnAceptar_Click(sender As Object, e As EventArgs) Handles btnAceptar.Click
        'Me.txtUsuario.Text = "Psicoweb"
        'Me.txtContrasena.Text = "@dm1n"
        'La variable de sesión que nos indica que el usuario esta logueado es usuarioID
        'La variable de sesión que nos indica de que nivel es: nivelID
        'Hay 3 niveles, 1=Master, 2=administrativo, 3=aplicador
        Dim entrar As String = String.Empty
        entrar = valida(Me.txtUsuario.Text.ToString, Me.txtContrasenia.Text.ToString)
        If entrar = String.Empty Then
            Dim enviaNivel As String = CType(Session.Item("nivelID"), String)
            Select Case (enviaNivel)
                Case 1
                    'nivel Master
                    Session.Add("Aplicador", False)
                    'Session.Add("Pruebas", 7)
                    Response.Redirect("~/Administracion/Default.aspx")
                Case 2
                    'nivel Administrativo
                    Session.Add("Aplicador", False)
                    Response.Redirect("Default.aspx")
                Case 3, 4
                    'nivel Aplicador
                    Response.Redirect("~/expediente/Instrucciones.aspx")
            End Select
        Else
            Session.Clear()
            Me.lblmensajes.Visible = True
            Me.LblMensajes.Text = entrar
        End If
    End Sub

    'Protected Sub ckbVer_CheckedChanged(sender As Object, e As EventArgs)
    '    'Me.lblpass.Text = Me.txtContrasenia.Text.ToString
    '    If Me.ckbVer.Checked = True Then
    '        Me.txtContrasenia.TextMode = TextBoxMode.SingleLine
    '    Else
    '        Me.txtContrasenia.TextMode = TextBoxMode.Password
    '        'Me.txtContrasenia.Text = Me.lblpass.Text.ToString
    '    End If
    'End Sub

End Class
