# coding: utf-8

import logging
import requests

from org.universolibre.EasyDev import XLORequests, Response
from easydev import comun
from easydev.setting import LOG, NAME_EXT, PY2, TIMEOUT

if PY2:
    str = unicode

log = logging.getLogger(NAME_EXT)


class LORequests(XLORequests):

    def __init__(self):
        pass

    @comun.catch_exception
    def requests(self, data):
        args = comun.to_dict(data.Args)
        info = comun.to_dict(data.Data)
        if data.Method == 'get':
            params = comun.to_dict(data.Params)
            r = requests.get(data.Url, params, **args)
        elif data.Method == 'post':
            json = comun.to_dict(data.Json)
            r = requests.post(data.Url, info, json, **args)
        elif data.Method == 'options':
            r = requests.options(data.Url, **args)
        elif data.Method == 'head':
            r = requests.head(data.Url, **args)
        elif data.Method == 'put':
            r = requests.put(data.Url, info, **args)
        elif data.Method == 'patch':
            r = requests.patch(data.Url, info, **args)
        elif data.Method == 'delete':
            r = requests.delete(data.Url, **args)

        res = Response()
        res.StatusCode = r.status_code
        res.Url = r.url
        res.Encoding = str(r.encoding)
        res.Headers = str(r.headers)
        res.Cookies = str(r.cookies)
        res.Text = r.text
        try:
            res.Json = str(r.json())
        except ValueError:
            res.Json = ''
        res.Content = r.content
        return res

