Imports CommandLine

Public Class CmdOption

#Region "SMTPServer"
    <[Option]("SMTPServer", Required:=True, HelpText:="SMTP server address")>
    Public Property SMTPServer() As String
        Get
            Return _SMTPServer
        End Get
        Set(ByVal value As String)
            _SMTPServer = value
        End Set
    End Property

    Private _SMTPServer As String
#End Region

#Region "SMTPPort"
    <[Option]("SMTPPort", Required:=True, HelpText:="SMTP server port")>
    Public Property SMTPPort() As Integer
        Get
            Return _SMTPPort
        End Get
        Set(ByVal value As Integer)
            _SMTPPort = value
        End Set
    End Property

    Private _SMTPPort As Integer
#End Region

#Region "SMTPSecurity"
    <[Option]("SMTPSecurity", Required:=True, HelpText:="SMTP server security")>
    Public Property SMTPSecurity() As Integer
        Get
            Return _SMTPSecurity
        End Get
        Set(ByVal value As Integer)
            _SMTPSecurity = value
        End Set
    End Property

    Private _SMTPSecurity As Integer
#End Region

#Region "SMTPAuth"
    <[Option]("SMTPAuth", Required:=True, HelpText:="SMTP server authentication")>
    Public Property SMTPAuth() As Integer
        Get
            Return _SMTPAuth
        End Get
        Set(ByVal value As Integer)
            _SMTPAuth = value
        End Set
    End Property

    Private _SMTPAuth As Integer
#End Region

#Region "SMTPUser"
    <[Option]("SMTPUser", HelpText:="SMTP username")>
    Public Property SMTPUser() As String
        Get
            Return _SMTPUser
        End Get
        Set
            _SMTPUser = Value
        End Set
    End Property

    Private _SMTPUser As String
#End Region

#Region "SMTPPwd"
    <[Option]("SMTPPwd", HelpText:="SMTP password")>
    Public Property SMTPPwd() As String
        Get
            Return _SMTPPwd
        End Get
        Set(ByVal value As String)
            _SMTPPwd = value
        End Set
    End Property

    Private _SMTPPwd As String
#End Region

#Region "From"
    <[Option]("From", Required:=True, HelpText:="Sender")>
    Public Property From() As String
        Get
            Return _From
        End Get
        Set(ByVal value As String)
            _From = value
        End Set
    End Property

    Private _From As String
#End Region

#Region "To"
    <[Option]("To", Required:=True, Separator:=";", HelpText:="Receiver")>
    Public Property [To]() As IEnumerable(Of String)
        Get
            Return _To
        End Get
        Set(ByVal value As IEnumerable(Of String))
            _To = value
        End Set
    End Property

    Private _To As IEnumerable(Of String)
#End Region

#Region "CC"
    <[Option]("CC", Separator:=";", HelpText:="CC")>
    Public Property CC() As IEnumerable(Of String)
        Get
            Return _CC
        End Get
        Set(ByVal value As IEnumerable(Of String))
            _CC = value
        End Set
    End Property

    Private _CC As IEnumerable(Of String)
#End Region

#Region "BCC"
    <[Option]("BCC", Separator:=";", HelpText:="BCC")>
    Public Property BCC() As IEnumerable(Of String)
        Get
            Return _BCC
        End Get
        Set(ByVal value As IEnumerable(Of String))
            _BCC = value
        End Set
    End Property

    Private _BCC As IEnumerable(Of String)
#End Region

#Region "Subject"
    <[Option]("Subject", Required:=True, HelpText:="Mail Subject")>
    Public Property Subject() As String
        Get
            Return _Subject
        End Get
        Set
            _Subject = Value
        End Set
    End Property

    Private _Subject As String
#End Region

#Region "Body"
    <[Option]("Body", Required:=True, HelpText:="Mail Body")>
    Public Property Body() As String
        Get
            Return _Body
        End Get
        Set
            _Body = Value
        End Set
    End Property

    Private _Body As String
#End Region

#Region "Attachment"
    <[Option]("Attachment", Separator:=";", HelpText:="Attachment files")>
    Public Property Attachment() As IEnumerable(Of String)
        Get
            Return _Attachment
        End Get
        Set(ByVal value As IEnumerable(Of String))
            _Attachment = value
        End Set
    End Property

    Private _Attachment As IEnumerable(Of String)
#End Region

#Region "Urgent"
    <[Option]("Urgent", HelpText:="Urgent flag")>
    Public Property Urgent() As Integer
        Get
            Return _Urgent
        End Get
        Set(ByVal value As Integer)
            _Urgent = value
        End Set
    End Property

    Private _Urgent As Integer
#End Region

#Region "Log"
    <[Option]("Log", HelpText:="Debug Log")>
    Public Property Log() As String
        Get
            Return _Log
        End Get
        Set(ByVal value As String)
            _Log = value
        End Set
    End Property

    Private _Log As String
#End Region

End Class
