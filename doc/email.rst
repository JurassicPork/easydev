Send mail
=========

Send mail and wait result
-------------------------

.. code-block:: vbnet

    Sub SendMail()
        util = createUnoService("org.universolibre.EasyDev")

        server = createUnoStruct("org.universolibre.EasyDev.SmtpServer")
        message = createUnoStruct("org.universolibre.EasyDev.EmailMessage")

        server.Name = "smtp.gmail.com"
        server.User = "hipatia.blades@gmail.com"
        server.Password = "supersecret"
        server.Ssl = True

        temp = "Dear Madame: $name\n\nBest regards from $country"
        data = Array( _
            Array("name", "Teresa"), _
            Array("country", "México"), _
        )
        body = util.render(temp, data)

        message.To = "public@mauriciobaeza.net"
        message.Subject = "Email test"
        message.Body = body

        'Send mail and wait response
        result = util.sendMail(server, message)
        If result Then
            MsgBox "Send mail"
        Else
            MsgBox "Error send mail"
        End If
    End Sub


Send mail and not wait
----------------------

.. code-block:: vbnet

    Sub SendMailNoWait()
        util = createUnoService("org.universolibre.EasyDev")

        server = createUnoStruct("org.universolibre.EasyDev.SmtpServer")
        message = createUnoStruct("org.universolibre.EasyDev.EmailMessage")

        server.Name = "smtp.gmail.com"
        server.User = "hipatia.blades@gmail.com"
        server.Password = "supersecret"
        server.Ssl = True
        server.Thread = True    'Send in other thread

        temp = "Dear Madame: $name\n\nBest regards from $country"
        data = Array( _
            Array("name", "Teresa"), _
            Array("country", "México"), _
        )
        body = util.render(temp, data)

        message.To = "public@mauriciobaeza.net"
        message.Subject = "Email test"
        message.Body = body

        'Send mail and NO wait response
        util.sendMail(server, message)
        MsgBox "Send mail"

    End Sub


Send mail with CC and BCC
-------------------------

.. code-block:: vbnet

    Sub SendMailNoWaitAndCCBCC()
        util = createUnoService("org.universolibre.EasyDev")

        server = createUnoStruct("org.universolibre.EasyDev.SmtpServer")
        message = createUnoStruct("org.universolibre.EasyDev.EmailMessage")

        server.Name = "smtp.gmail.com"
        server.User = "hipatia.blades@gmail.com"
        server.Password = "supersecret"
        server.Ssl = True
        server.Thread = True    'Send in other thread

        temp = "Dear Madame: $name\n\nBest regards from $country"
        data = Array( _
            Array("name", "Teresa"), _
            Array("country", "México"), _
        )
        body = util.render(temp, data)

        message.To = "public@mauriciobaeza.net"
        message.Subject = "Email test"
        message.Body = body
        message.Cc = "other1@correo.net,other2@correo.net"
        message.Bcc = "other3@correo.net,other4@correo.net"

        'Send mail and NO wait response
        util.sendMail(server, message)
        MsgBox "Send mail"

    End Sub


Send mail with attachment
-------------------------

.. code-block:: vbnet

    Sub SendMailWithAttachment()
        util = createUnoService("org.universolibre.EasyDev")

        server = createUnoStruct("org.universolibre.EasyDev.SmtpServer")
        message = createUnoStruct("org.universolibre.EasyDev.EmailMessage")

        server.Name = "smtp.gmail.com"
        server.User = "hipatia.blades@gmail.com"
        server.Password = "supersecret"
        server.Ssl = True
        server.Thread = True    'Send in other thread

        temp = "Dear Madame: $name\n\nBest regards from $country"
        data = Array( _
            Array("name", "Teresa"), _
            Array("country", "México"), _
        )
        body = util.render(temp, data)

        files = Array("/home/USER/Documents/pruebas.cer", "/home/USER/Documents/pruebas.key")

        message.To = "public@mauriciobaeza.net"
        message.Subject = "Email test"
        message.Body = body
        message.Files = files

        'Send mail and NO wait response
        util.sendMail(server, message)
        MsgBox "Send mail"

    End Sub


Send mail and save
------------------

If you use a path in Thunderbird, you can see mail in folder into Thunderbird

.. CAUTION::
   Use separate folder into tree folders in Thunderbird!

.. code-block:: vbnet

    Sub SendMailAndSave()
        util = createUnoService("org.universolibre.EasyDev")

        server = createUnoStruct("org.universolibre.EasyDev.SmtpServer")
        message = createUnoStruct("org.universolibre.EasyDev.EmailMessage")

        server.Name = "smtp.gmail.com"
        server.User = "hipatia.blades@gmail.com"
        server.Password = "supersecret"
        server.Ssl = True
        server.Thread = True    'Send in other thread
        server.PathSave = "/home/USER/.thunderbird/cwfln0bi.default/Mail/Local Folders/Sent"

        temp = "Dear Madame: $name\n\nBest regards from $country"
        data = Array( _
            Array("name", "Teresa"), _
            Array("country", "México"), _
        )
        body = util.render(temp, data)

        files = Array("/home/USER/Documents/pruebas.cer")

        message.To = "public@mauriciobaeza.net"
        message.Subject = "Email test"
        message.Body = body
        message.Files = files
        message.Save = True

        'Send mail and NO wait response
        util.sendMail(server, message)
        MsgBox "Send mail"

    End Sub

.. image:: images/img008.png
    :width: 800px
    :align: center
