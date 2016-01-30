Requests
========

`Requests`_ is the best library for HTTP in Python


Get
---

.. code-block:: vbnet

    Sub Get
        util = createUnoService("org.universolibre.EasyDev")
        data = createUnoStruct("org.universolibre.EasyDev.Requests")


        data.Method = "get"
        data.Url = "https://api.vaultoro.com/latest"
        args = Array( _
            Array("verify", False) _
        )
        data.Args = args

        response = util.requests(data)
        util.msgbox(response.Text)
    End Sub


First, get IP public, and then, get location IP

.. code-block:: vbnet

    Sub GetIPLocation()
        util = createUnoService("org.universolibre.EasyDev")
        data = createUnoStruct("org.universolibre.EasyDev.Requests")

        data.Method = "get"
        data.Url = "http://api.ipify.org"
        response = util.requests(data)
        ip_public = response.Text

        data.Url = "http://freegeoip.net/csv/" & ip_public
        response = util.requests(data)
        util.msgbox(response.Text)

    End Sub


Use Auth

.. code-block:: vbnet

    util = createUnoService("org.universolibre.EasyDev")
    data = createUnoStruct("org.universolibre.EasyDev.Requests")

    data.Method = "get"
    data.Url = "https://api.github.com/user"
    data.Args = Array( _
        Array("auth", Array("mauriciobaeza", "supersecret")) _
    )

    response = util.requests(data)

    util.msgbox(response.StatusCode)
    util.msgbox(response.Text)

You can test request in: `<http://httpbin.org/>`_


.. _Requests: http://docs.python-requests.org/en/latest/
