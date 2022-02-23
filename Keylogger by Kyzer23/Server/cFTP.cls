VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDOmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' para la conexión a internet
Private Declare Function InternetGetConnectedState _
    Lib "wininet.dll" ( _
    ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long

Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
Private Const INTERNET_RAS_INSTALLED As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

' variables locales
Private mServidor As String
Private mPara As String
Private mDe As String
Private mAsunto As String
Private mMensaje As String
Private mAdjunto As String
Private mPuerto As Variant
Private mUsuario As String
Private mPassword As String
Private mUseAuntentificacion As Boolean
Private mSSL As Boolean

Public Event Error(Descripcion As String, Numero As Variant)
Public Event EnvioCompleto()

Function Enviar_Backup() As Boolean
    
    ' Variable de objeto Cdo.Message
    Dim oCDO As Object
          
    ' chequea si hay conexión
    If InternetGetConnectedState(0&, 0&) = False Then
       RaiseEvent Error("No se puede enviar el correo. " & _
                        "Verificar la conexión a internet si está disponible", 0)
       Exit Function
    End If
    
    
    
    ' chequea que el puerto sea un número, o que no esté vacío
    If Not IsNumeric(puerto) Then
       RaiseEvent Error("No se ha indicado el puerto del servidor", 0)
       Exit Function
    End If
    
    ' Crea un Nuevo objeto CDO.Message
    Set oCDO = CreateObject("CDO.Message")
    
    ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
      del servidor o su dirección IP )
    oCDO.Configuration.Fields( _
    "http://schemas.microsoft.com/cdo/configuration/smtpserver") = mServidor
    
    oCDO.Configuration.Fields( _
    "http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    
    ' Puerto. Por defecto se usa el puerto 25, _
     en el caso de Gmail se usa el puerto 465
    
    oCDO.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = mPuerto

    
    ' Indica el tipo de autentificación con el servidor de correo _
     El valor 0 no requiere autentificarse, el valor 1 es con autentificación
    oCDO.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & _
                "configuration/smtpauthenticate") = Abs(mUseAuntentificacion)
    
    ' Tiempo máximo de espera en segundos para la conexión
    oCDO.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10

    ' Configura las opciones para el login en el SMTP
    If mUseAuntentificacion Then

    ' Id de usuario del servidor Smtp ( en el caso de gmail, _
     debe ser la dirección de correro mas el @gmail.com )
    oCDO.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendusername") = mUsuario

    ' Password de la cuenta
    oCDO.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = mPassword

    ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
    oCDO.Configuration.Fields.Item _
        ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = mSSL
    
    End If
    
    ' Estructura del mail
    '''''''''''''''''''''''''''''''''''''''''''''''
    
    ' Dirección del Destinatario
    
    oCDO.To = mPara
    
    ' Dirección del remitente
    oCDO.From = mDe
    
    ' Asunto del mensaje
    oCDO.Subject = mAsunto
    
    ' Cuerpo del mensaje
    oCDO.TextBody = mMensaje
    
    'Ruta del archivo adjunto
    If mAdjunto <> "" Then
        If Len(Dir(mAdjunto)) = 0 Then
            ' ..error
            RaiseEvent Error("No se ha encontrado el archivo en la siguiente ruta: ", 0)
            Exit Function
        Else
            ' ..lo agrega
            oCDO.AddAttachment (mAdjunto)
        End If
    End If
    
    ' Actualiza los datos antes de enviar
    oCDO.Configuration.Fields.Update
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    ' Envía el email
    oCDO.Send
    Screen.MousePointer = 0
    ' .. si no hubo error
    If Err.Number = 0 Then
       Enviar_Backup = True
       RaiseEvent EnvioCompleto
    
    ElseIf Err.Number = -2147220973 Then
       RaiseEvent Error("Posible error : nombre del Servidor " & _
                        "incorrecto o número de puerto incorrecto", Err.Number)
    ElseIf Err.Number = -2147220975 Then
       RaiseEvent Error("Posible error : error en la el nombre de usuario, " & _
                                                "o en el password ", Err.Number)
    Else
       RaiseEvent Error(Err.Description, Err.Number)
    End If

    ' Descarga la referencia
    If Not oCDO Is Nothing Then
        Set oCDO = Nothing
    End If
    
    Err.Clear
    
    Screen.MousePointer = vbNormal
End Function

' propiedades
'''''''''''''''''''''
Property Get servidor() As String
    servidor = mServidor
End Property
Property Let servidor(value As String)
    mServidor = value
End Property


Property Get para() As String
    para = mPara
End Property
Property Let para(value As String)
    mPara = value
End Property


Property Get de() As String
    de = mDe
End Property
Property Let de(value As String)
    mDe = value
End Property


Property Get Asunto() As String
    Asunto = mAsunto
End Property
Property Let Asunto(value As String)
    mAsunto = value
End Property


Property Get Mensaje() As String
    Mensaje = mMensaje
End Property
Property Let Mensaje(value As String)
    mMensaje = value
End Property


Property Get Adjunto() As String
    Adjunto = mAdjunto
End Property
Property Let Adjunto(value As String)
    mAdjunto = value
End Property


Property Get puerto() As Variant
    puerto = mPuerto
End Property
Property Let puerto(value As Variant)
    mPuerto = value
End Property


Property Get Usuario() As String
    Usuario = mUsuario
End Property
Property Let Usuario(value As String)
    mUsuario = value
End Property


Property Get PassWord() As String
    PassWord = mPassword
End Property
Property Let PassWord(value As String)
    mPassword = value
End Property


Property Get UseAuntentificacion() As Boolean
    UseAuntentificacion = mUseAuntentificacion
End Property
Property Let UseAuntentificacion(value As Boolean)
    mUseAuntentificacion = value
End Property


Property Get ssl() As Boolean
    ssl = mSSL
End Property
Property Let ssl(value As Boolean)
    mSSL = value
End Property



