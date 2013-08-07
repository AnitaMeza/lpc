''  Fecha de creación:   16/01/2013 12:33:00 p.m.
''  Descripción de la funcionalidad:   Clase BOeel_Preguntas funcionará para gestionar las opciones de CRUD y funciones auxiliares para persistir la información en la BD o funciones que ayuden a manipular la información
''  Autor:   Rafael Ramírez Luna
''  ******************************************************************************************************************
''  Fecha de la última modificación:   
''  Autor de la última modificación:   
''  Descripción de la última modificación:   
''  ******************************************************************************************************************
''  Fecha de la última modificación:   
''  Autor de la última modificación:   
''  Descripción de la última modificación:   
''  ******************************************************************************************************************

Option Explicit On
Option Strict On

Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
''Espacio de nombres para utilizar los Business Objects de BOeel_Preguntas
Imports BO

Namespace DAL

    Public Class DALeel_Preguntas

#Region "Propiedades"

        Private _conexion As String = String.Empty

        Public Property Conexion As String
            Get
                Return _conexion
            End Get
            Set(value As String)
                _conexion = value
            End Set
        End Property

        Private _mensaje As String = String.Empty

        Public Property Mensaje As String
            Get
                Return _mensaje
            End Get
            Set(value As String)
                _mensaje = value
            End Set
        End Property

#End Region

#Region "Constructores"

        ''' <summary>
        ''' Constructor por default
        ''' </summary>
        Public Sub New()

        End Sub

        ''' <summary>
        ''' Constructor para poder inicializar la cadena de conexión
        ''' </summary>
        ''' <param name="Conexion"></param>
        ''' <remarks></remarks>
        Public Sub New(conexion As String)
            _conexion = conexion
        End Sub
#End Region

#Region "Métodos comunes"

        '''<summary>
        ''' Obtiene todos los objetos BOeel_Preguntas de la base de datos
        '''</summary>
        ''' <returns>Retorna una lista de objetos BOeel_Preguntas que se encuentren en la BD</returns>
        Public Function obtenerTodo() As List(Of BOeel_Preguntas)

            Dim sentencia As New czSentenciasSinTransaccion()
            Dim cmd As New SqlCommand()

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasObtenerTodo", cmd, _conexion)

            _mensaje = cmd.Parameters("@MensajeError").Value.ToString()

            If (_mensaje <> String.empty) Then
                Return Nothing
            Else

                Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)
                For Each drw As DataRow In dt.Rows
                    Dim dr As DataRow = drw
                    Dim temp As BOeel_Preguntas = obtenerObjeto(dr)
                    listaBOeel_Preguntas.Add(temp)
                Next

            End If

            Return listaBOeel_Preguntas

        End Function

        ''' <summary>
        ''' Obtiene un objeto BOeel_Preguntas recuperado con las llaves primarias pasadas como parámetros del método
        ''' </summary>
        ''' <remarks>Regresa un objeto BOeel_Preguntas en caso de ser encontrado, en caso contrario nulo</remarks>
        Public Function obtenerPorId(ByVal Id As Int32) As BOeel_Preguntas
            Dim cmd As New SqlCommand()
            'Asignación de variables que se pasan por el método al commando de sql para obtener los registros
            cmd.Parameters.AddWithValue("@Id", Id)
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output


            Dim sentencia As New czSentenciasSinTransaccion()
            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasObtenerPorId", cmd, _conexion)

            If (dt.Rows.Count > 0) Then
                Dim dr As DataRow = dt.Rows(0)
                'Asignación del objeto obtenido a la entidad
                Dim temp As BOeel_Preguntas = obtenerObjeto(dr)
                Return temp
            End If
            'Si no se encontró ninguno, retornaremos nothing
            Return Nothing
        End Function


        '''<summary>
        ''' Actualiza una entidad BOeel_Preguntas se le proporciona el objeto con la llave primaria en el mismo objeto
        '''</summary>
        ''' <param name="entidad">Entidad a actualizar, las llaves primarias deben existir en la BD</param>
        ''' <returns>retorna el número de registros actualizados</returns>
        Public Function actualizar(entidad As BOeel_Preguntas) As Boolean
            Dim sentencia As New czSentenciasSinTransaccion()
            'Cargamos todos los parámetros de SQL desde la entidad BOeel_Preguntas
            Dim cmd As New SqlCommand()
            cmd.CommandText = "usp_eel_PreguntasActualizar"
            cmd.Parameters.AddWithValue("@Id", entidad.Id)
            cmd.Parameters.AddWithValue("@Descripcion", entidad.Descripcion)
            cmd.Parameters.AddWithValue("@Duracion", entidad.Duracion)
            cmd.Parameters.AddWithValue("@Orden", entidad.Orden)
            cmd.Parameters.AddWithValue("@IdEntrevista", entidad.IdEntrevista)
            cmd.Parameters.AddWithValue("@Calificacion", entidad.Calificacion)
            cmd.Parameters.AddWithValue("@DuracionPregunta", entidad.DuracionPregunta)
            cmd.Parameters.AddWithValue("@IdVideo", entidad.IdVideo)
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output


            Return (sentencia.ejecutarProcedimientoAlmacenado(cmd, _conexion) > 0)
        End Function

        ''' <summary>
        ''' Elimina un objeto BOeel_Preguntas, únicamente recibiendo la o las llaves primarias
        ''' </summary>
        Public Function eliminar(ByVal Id As Int32) As Boolean
            'Creamos el objeto que se encargará de ejcutar sentencias
            Dim sentencia As New czSentenciasSinTransaccion()
            'Cargamos todos los parámetros de SQL desde la entidad BOeel_Preguntas
            Dim cmd As New SqlCommand()
            'Configuramos el objeto para eliminar el registro
            cmd.CommandText = "usp_eel_PreguntasEliminar"
            'Asignación de variables que se pasan por el método al commando de sql para obtener los registros
            cmd.Parameters.AddWithValue("@Id", Id)
            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            If (Not sentencia.ejecutarProcedimientoAlmacenado(cmd, _conexion)) Then
                _mensaje = cmd.Parameters("@MensajeError").Value.ToString()
                Return False
            End If

            Return True
        End Function


        ''' <summary>
        ''' Elimina un lista de objetos BOeel_Preguntas, cada objeto BOeel_Preguntas debe tener asignado los valores de las llaves foráneas únicamente
        ''' </summary>
        ''' <param name="listaBOeel_Preguntas">Cada objeto BOeel_Preguntas debe tener asignado los valores de las llaves foráneas, para poder eliminarlo</param>
        ''' <returns></returns>
        Public Function eliminarLista(listaBOeel_Preguntas As List(Of BOeel_Preguntas)) As Boolean
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Creamos el objeto que se encargará de ejcutar sentencias
            Dim sentencia As New czSentenciasSinTransaccion()
            'Configuramos el objeto para eliminar el registro
            cmd.CommandText = "usp_eel_PreguntasEliminar"


            Using tran As New TransactionScope


                For Each entidad As BOeel_Preguntas In listaBOeel_Preguntas
                    'Asignación de variables que se pasan por el método al commando de sql para eliminar los objetos de BOeel_Preguntas
                    cmd.Parameters.AddWithValue("@Id", entidad.Id)

                    'Parámetro para obtener la descripción del error en caso de que exista
                    cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

                    If (Not sentencia.ejecutarProcedimientoAlmacenado(cmd, _conexion)) Then
                        _mensaje = cmd.Parameters("@MensajeError").Value.ToString()
                        Return False
                    End If

                    'borramos los parámetros para que en la otra iteracción se configure con los nuevos parámetros
                    cmd.Parameters.Clear()
                Next

                tran.Complete()

            End Using

            Return True
        End Function


        ''' <summary>
        ''' Inserta un objeto BOeel_Preguntas a la Base de Datos
        '''</summary>
        ''' <param name="entidad">Objeto BOeel_Preguntas a insertar en la BD</param>
        ''' <returns>Retorna y true en caso de ser insertado, en caso contrario false</returns>
        Public Function insertar(entidad As BOeel_Preguntas) As Boolean
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Creamos el objeto que se encargará de ejcutar sentencias
            Dim sentencia As New czSentenciasSinTransaccion()
            'Configuramos el objeto para insertar el registro
            cmd.CommandText = "usp_eel_PreguntasInsertar"
            'Asignamos los parámetros de SQL desde la entidad que se pasó como parámetros

            'Parámetro para obtener el valor identidad del registro que se acaba de insertar
            cmd.Parameters.Add("@ReferenciaId", SqlDbType.Int).Direction = System.Data.ParameterDirection.Output
            cmd.Parameters.AddWithValue("@Descripcion", entidad.Descripcion)
            cmd.Parameters.AddWithValue("@Duracion", entidad.Duracion)
            cmd.Parameters.AddWithValue("@Orden", entidad.Orden)
            cmd.Parameters.AddWithValue("@IdEntrevista", entidad.IdEntrevista)
            cmd.Parameters.AddWithValue("@Calificacion", entidad.Calificacion)
            cmd.Parameters.AddWithValue("@DuracionPregunta", entidad.DuracionPregunta)
            cmd.Parameters.AddWithValue("@IdVideo", entidad.IdVideo)





            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output



            Dim resultado As Integer = sentencia.ejecutarProcedimientoAlmacenado(cmd, _conexion)
            If (resultado > 0) Then
                entidad.Id = CType(cmd.Parameters("@ReferenciaId").Value, Int32)
            End If

            Return (resultado > 0)
        End Function

        Public Function insertarLista(listaBOeel_Preguntas As List(Of BOeel_Preguntas)) As Boolean
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Creamos el objeto que se encargará de ejcutar sentencias
            Dim sentencia As New czSentenciasSinTransaccion()
            'Configuramos el objeto para eliminar el registro
            cmd.CommandText = "usp_eel_PreguntasEliminar"

            Using tran As New TransactionScope

                For Each entidad As BOeel_Preguntas In listaBOeel_Preguntas

                    cmd.Parameters.Add("@ReferenciaId", SqlDbType.Int).Direction = System.Data.ParameterDirection.Output
                    cmd.Parameters.AddWithValue("@Descripcion", entidad.Descripcion)
                    cmd.Parameters.AddWithValue("@Duracion", entidad.Duracion)
                    cmd.Parameters.AddWithValue("@Orden", entidad.Orden)
                    cmd.Parameters.AddWithValue("@IdEntrevista", entidad.IdEntrevista)
                    cmd.Parameters.AddWithValue("@Calificacion", entidad.Calificacion)
                    cmd.Parameters.AddWithValue("@DuracionPregunta", entidad.DuracionPregunta)
                    cmd.Parameters.AddWithValue("@IdVideo", entidad.IdVideo)


                    If (Not sentencia.ejecutarProcedimientoAlmacenado(cmd, _conexion)) Then
                        _mensaje = cmd.Parameters("@MensajeError").Value.ToString()
                        Return False
                    End If

                    entidad.Id = CType(cmd.Parameters("@ReferenciaId").Value, Int32)

                    'borramos los parámetros para que en la otra iteracción se configure con los nuevos parámetros
                    cmd.Parameters.Clear()
                Next
                tran.Complete()
            End Using

            Return True
        End Function
#End Region

#Region "Métodos auxiliares"

        ''' <summary>
        ''' Genera un objeto de tipo BOeel_Preguntas dependiendo del DataRow especificado en el parámetro
        ''' </summary>
        ''' <param name="dr">DataRow de donde se obtendrá los valores para generar el objeto BOeel_Preguntas</param>
        ''' <returns>Retorna una un objeto BOeel_Preguntas con los valores asignados del DataRow</returns>
        Private Function obtenerObjeto(dr As DataRow) As BOeel_Preguntas
            Dim temp As New BOeel_Preguntas()
            temp.Id = CType(dr("Id"), Int32)
            temp.Descripcion = CType(dr("Descripcion"), String)
            temp.Duracion = CType(dr("Duracion"), Int32)
            temp.Orden = CType(dr("Orden"), Int32)
            temp.IdEntrevista = CType(dr("IdEntrevista"), Int32)
            temp.Calificacion = CType(dr("Calificacion"), Decimal)
            temp.DuracionPregunta = CType(dr("DuracionPregunta"), Int32)
            If (dr("IdVideo") Is System.DBNull.Value) Then
                temp.IdVideo = Nothing
            Else
                temp.IdVideo = CType(dr("IdVideo"), String)

            End If

            Return temp
        End Function

        ''' <summary>
        ''' Obtiene una lista de objetos BOeel_Preguntas filtrado por la condición que es pasada como parámetro
        ''' </summary>
        ''' <param name="condicion">Condición con el cuál será filtrado <example>Nombre like '%Rafa%'</example><example>Sueldo >= 2000 And Gasto = 0 or Nombre Like '%rafa%' </example> este método pone el WHERE </param>
        ''' <returns>una lista de objetos BOeel_Preguntas</returns>
        ''' <remarks>rlr</remarks>
        Public Function obtenerListaWhere(condicion As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            If (condicion.Trim() <> String.Empty) Then
                condicion = "WHERE " & condicion
            End If
            Dim dt As DataTable = sentencia.obtenerDataTable(String.Format("SELECT * FROM eel_Preguntas {0};", condicion), _conexion)
            If (dt.Rows.Count > 0) Then
                For Each drw As DataRow In dt.Rows
                    Dim dr As DataRow = drw
                    listaBOeel_Preguntas.Add(obtenerObjeto(dr))
                Next
                'Retornamos la lista que se generó
                Return listaBOeel_Preguntas
            End If
            Return Nothing
        End Function
        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdMayorQue" de la tabla, parámetros necesarios: ByVal Id as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdMayorQue" (registros encontrados)</returns>
        Public Function obtenerPorIdMayorQue(ByVal Id As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Id", Id)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdMayorQue", cmd, _conexion)
            If (dt.Rows.Count > 0) Then
                For Each drw As DataRow In dt.Rows
                    listaBOeel_Preguntas.Add(obtenerObjeto(drw))
                Next
                'Retornamos la lista que se generó
                Return listaBOeel_Preguntas
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdMenorQue" de la tabla, parámetros necesarios: ByVal Id as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdMenorQue" (registros encontrados)</returns>
        Public Function obtenerPorIdMenorQue(ByVal Id As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Id", Id)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdMenorQue", cmd, _conexion)
            If (dt.Rows.Count > 0) Then
                For Each drw As DataRow In dt.Rows
                    listaBOeel_Preguntas.Add(obtenerObjeto(drw))
                Next
                'Retornamos la lista que se generó
                Return listaBOeel_Preguntas
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "Descripcion" de la tabla, parámetros necesarios: ByVal Descripcion as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "Descripcion" (registros encontrados)</returns>
        Public Function obtenerPorDescripcion(ByVal Descripcion As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Descripcion", Descripcion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDescripcion", cmd, _conexion)
            If (dt.Rows.Count > 0) Then
                For Each drw As DataRow In dt.Rows
                    listaBOeel_Preguntas.Add(obtenerObjeto(drw))
                Next
                'Retornamos la lista que se generó
                Return listaBOeel_Preguntas
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DescripcionContiene" de la tabla, parámetros necesarios: ByVal Descripcion as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DescripcionContiene" (registros encontrados)</returns>
        Public Function obtenerPorDescripcionContiene(ByVal Descripcion As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Descripcion", Descripcion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDescripcionContiene", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DescripcionIniciaCon" de la tabla, parámetros necesarios: ByVal Descripcion as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DescripcionIniciaCon" (registros encontrados)</returns>
        Public Function obtenerPorDescripcionIniciaCon(ByVal Descripcion As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Descripcion", Descripcion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDescripcionIniciaCon", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DescripcionTerminaCon" de la tabla, parámetros necesarios: ByVal Descripcion as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DescripcionTerminaCon" (registros encontrados)</returns>
        Public Function obtenerPorDescripcionTerminaCon(ByVal Descripcion As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Descripcion", Descripcion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDescripcionTerminaCon", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "Duracion" de la tabla, parámetros necesarios: ByVal Duracion as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "Duracion" (registros encontrados)</returns>
        Public Function obtenerPorDuracion(ByVal Duracion As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Duracion", Duracion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDuracion", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DuracionMayorQue" de la tabla, parámetros necesarios: ByVal Duracion as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DuracionMayorQue" (registros encontrados)</returns>
        Public Function obtenerPorDuracionMayorQue(ByVal Duracion As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Duracion", Duracion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDuracionMayorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DuracionMenorQue" de la tabla, parámetros necesarios: ByVal Duracion as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DuracionMenorQue" (registros encontrados)</returns>
        Public Function obtenerPorDuracionMenorQue(ByVal Duracion As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Duracion", Duracion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDuracionMenorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "Orden" de la tabla, parámetros necesarios: ByVal Orden as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "Orden" (registros encontrados)</returns>
        Public Function obtenerPorOrden(ByVal Orden As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Orden", Orden)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorOrden", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "OrdenMayorQue" de la tabla, parámetros necesarios: ByVal Orden as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "OrdenMayorQue" (registros encontrados)</returns>
        Public Function obtenerPorOrdenMayorQue(ByVal Orden As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Orden", Orden)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorOrdenMayorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "OrdenMenorQue" de la tabla, parámetros necesarios: ByVal Orden as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "OrdenMenorQue" (registros encontrados)</returns>
        Public Function obtenerPorOrdenMenorQue(ByVal Orden As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Orden", Orden)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorOrdenMenorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdEntrevista" de la tabla, parámetros necesarios: ByVal IdEntrevista as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdEntrevista" (registros encontrados)</returns>
        Public Function obtenerPorIdEntrevista(ByVal IdEntrevista As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@IdEntrevista", IdEntrevista)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdEntrevista", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdEntrevistaMayorQue" de la tabla, parámetros necesarios: ByVal IdEntrevista as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdEntrevistaMayorQue" (registros encontrados)</returns>
        Public Function obtenerPorIdEntrevistaMayorQue(ByVal IdEntrevista As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@IdEntrevista", IdEntrevista)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdEntrevistaMayorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdEntrevistaMenorQue" de la tabla, parámetros necesarios: ByVal IdEntrevista as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdEntrevistaMenorQue" (registros encontrados)</returns>
        Public Function obtenerPorIdEntrevistaMenorQue(ByVal IdEntrevista As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@IdEntrevista", IdEntrevista)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdEntrevistaMenorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "Calificacion" de la tabla, parámetros necesarios: ByVal Calificacion as Decimal
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "Calificacion" (registros encontrados)</returns>
        Public Function obtenerPorCalificacion(ByVal Calificacion As Decimal) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Calificacion", Calificacion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorCalificacion", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "CalificacionMayorQue" de la tabla, parámetros necesarios: ByVal Calificacion as Decimal
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "CalificacionMayorQue" (registros encontrados)</returns>
        Public Function obtenerPorCalificacionMayorQue(ByVal Calificacion As Decimal) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Calificacion", Calificacion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorCalificacionMayorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "CalificacionMenorQue" de la tabla, parámetros necesarios: ByVal Calificacion as Decimal
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "CalificacionMenorQue" (registros encontrados)</returns>
        Public Function obtenerPorCalificacionMenorQue(ByVal Calificacion As Decimal) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@Calificacion", Calificacion)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorCalificacionMenorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DuracionPregunta" de la tabla, parámetros necesarios: ByVal DuracionPregunta as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DuracionPregunta" (registros encontrados)</returns>
        Public Function obtenerPorDuracionPregunta(ByVal DuracionPregunta As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@DuracionPregunta", DuracionPregunta)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDuracionPregunta", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DuracionPreguntaMayorQue" de la tabla, parámetros necesarios: ByVal DuracionPregunta as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DuracionPreguntaMayorQue" (registros encontrados)</returns>
        Public Function obtenerPorDuracionPreguntaMayorQue(ByVal DuracionPregunta As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@DuracionPregunta", DuracionPregunta)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDuracionPreguntaMayorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "DuracionPreguntaMenorQue" de la tabla, parámetros necesarios: ByVal DuracionPregunta as Int32
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "DuracionPreguntaMenorQue" (registros encontrados)</returns>
        Public Function obtenerPorDuracionPreguntaMenorQue(ByVal DuracionPregunta As Int32) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@DuracionPregunta", DuracionPregunta)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorDuracionPreguntaMenorQue", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdVideo" de la tabla, parámetros necesarios: ByVal IdVideo as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdVideo" (registros encontrados)</returns>
        Public Function obtenerPorIdVideo(ByVal IdVideo As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@IdVideo", IdVideo)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdVideo", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdVideoContiene" de la tabla, parámetros necesarios: ByVal IdVideo as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdVideoContiene" (registros encontrados)</returns>
        Public Function obtenerPorIdVideoContiene(ByVal IdVideo As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@IdVideo", IdVideo)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdVideoContiene", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdVideoIniciaCon" de la tabla, parámetros necesarios: ByVal IdVideo as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdVideoIniciaCon" (registros encontrados)</returns>
        Public Function obtenerPorIdVideoIniciaCon(ByVal IdVideo As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@IdVideo", IdVideo)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdVideoIniciaCon", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function

        ''' <summary>
        ''' Obtiene registros buscados por el campo "IdVideoTerminaCon" de la tabla, parámetros necesarios: ByVal IdVideo as string
        ''' </summary>
        ''' <returns>Retorna una lista con los objetos BOeel_Preguntas encontrados por el campo: "IdVideoTerminaCon" (registros encontrados)</returns>
        Public Function obtenerPorIdVideoTerminaCon(ByVal IdVideo As String) As List(Of BOeel_Preguntas)
            Dim listaBOeel_Preguntas As New List(Of BOeel_Preguntas)()
            Dim sentencia As New czSentenciasSinTransaccion()
            'Creamos el Objeto SqlCommand para ejecutar las sentencias
            Dim cmd As New SqlCommand()
            'Parámetro de salida para obtener el número de error en caso de que exista
            cmd.Parameters.Add("@noError", SqlDbType.Int, 4).Direction = System.Data.ParameterDirection.Output

            'Parámetro para obtener la descripción del error en caso de que exista
            cmd.Parameters.Add("@MensajeError", SqlDbType.NVarChar, 4000).Direction = System.Data.ParameterDirection.Output

            cmd.Parameters.AddWithValue("@IdVideo", IdVideo)


            Dim dt As DataTable = sentencia.ejecutarProcedimientoAlmacenado("usp_eel_PreguntasobtenerPorIdVideoTerminaCon", cmd, _conexion)

            For Each drw As DataRow In dt.Rows
                listaBOeel_Preguntas.Add(obtenerObjeto(drw))
            Next
            'Retornamos la lista que se generó
            Return listaBOeel_Preguntas
        End Function


#End Region

    End Class

End Namespace
