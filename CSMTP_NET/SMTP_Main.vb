
Imports System.IO
Imports MailKit
Imports MailKit.Net.Smtp
Imports MimeKit

Module SMTP_Main

    Sub Main(ByVal szArgs() As String)

        ' Intestazione
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("***************************************")
        Console.WriteLine("*              CSMTP.NET              *")
        Console.WriteLine("***************************************")
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("")

        ' Controllo parametri linea di comando
        If szArgs.Length < 1 Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("[!] - Nessun parametro passato")
            Console.ForegroundColor = ConsoleColor.White
            Console.WriteLine("")
            Exit Sub
        End If

        ' Parsing parametri
        Dim pResult = CommandLine.Parser.Default.ParseArguments(Of CmdOption)(szArgs)

        ' Errore nel parsing
        If pResult.Tag = CommandLine.ParserResultType.NotParsed Then
            Exit Sub
        End If

        Dim pCommandResult As CommandLine.Parsed(Of CmdOption) = pResult

        ' Impostazione messaggio
        Dim pMessage = New MimeMessage()
        Dim dwIndex As Integer

        ' From
        pMessage.From.Add(New MailboxAddress(pCommandResult.Value.From, pCommandResult.Value.From))

        ' To
        For dwIndex = 0 To pCommandResult.Value.To.Count - 1
            pMessage.To.Add(New MailboxAddress(pCommandResult.Value.To(dwIndex).ToString, pCommandResult.Value.To(dwIndex).ToString))
        Next

        ' CC
        For dwIndex = 0 To pCommandResult.Value.CC.Count - 1
            pMessage.Cc.Add(New MailboxAddress(pCommandResult.Value.CC(dwIndex).ToString, pCommandResult.Value.CC(dwIndex).ToString))
        Next

        ' BCC
        For dwIndex = 0 To pCommandResult.Value.BCC.Count - 1
            pMessage.Bcc.Add(New MailboxAddress(pCommandResult.Value.BCC(dwIndex).ToString, pCommandResult.Value.BCC(dwIndex).ToString))
        Next

        ' Subject
        pMessage.Subject = pCommandResult.Value.Subject

        ' Body
        Dim pBody = New TextPart("plain")
        pBody.Text = pCommandResult.Value.Body

        ' Attachment
        If pCommandResult.Value.Attachment.Count > 0 Then

            Dim pMultipart = New Multipart("mixed")
            pMultipart.Add(pBody)

            For dwIndex = 0 To pCommandResult.Value.Attachment.Count - 1

                If Not File.Exists(pCommandResult.Value.Attachment(dwIndex)) Then
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("[!] - " + pCommandResult.Value.Attachment(dwIndex) + " non trovato")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine("")
                    Exit Sub
                End If

                Dim pAttachment = New MimePart(Path.GetFileName(pCommandResult.Value.Attachment(dwIndex)), Path.GetExtension(pCommandResult.Value.Attachment(dwIndex)))
                pAttachment.ContentObject = New ContentObject(File.OpenRead(pCommandResult.Value.Attachment(dwIndex)), ContentEncoding.Default)
                pAttachment.ContentDisposition = New ContentDisposition(ContentDisposition.Attachment)
                pAttachment.ContentTransferEncoding = ContentEncoding.Base64
                pMultipart.Add(pAttachment)

            Next

            pMessage.Body = pMultipart
        Else
            pMessage.Body = pBody
        End If

        Dim pClient As SmtpClient

        ' Init SmtpClient
        If pCommandResult.Value.Log <> Nothing Then
            pClient = New SmtpClient(New ProtocolLogger(pCommandResult.Value.Log))
        Else
            pClient = New SmtpClient()
        End If

        Try
            ' Connect
            pClient.Connect(pCommandResult.Value.SMTPServer, pCommandResult.Value.SMTPPort, IIf(pCommandResult.Value.SMTPSecurity = 1, True, False))

            ' Authenticate
            If pCommandResult.Value.SMTPAuth = 1 Then
                pClient.Authenticate(pCommandResult.Value.SMTPUser, pCommandResult.Value.SMTPPwd)
            End If

            ' Send
            pClient.Send(pMessage)

            ' disconnect
            pClient.Disconnect(True)
        Catch ex As Exception
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("[!] - " + ex.Message)
            Console.ForegroundColor = ConsoleColor.White
            Console.WriteLine("")
            Exit Sub
        End Try

    End Sub

End Module
