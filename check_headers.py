# -*- coding: utf-8 -*-
import urllib.request
import json

try:
    req = urllib.request.urlopen('http://localhost:8765/', timeout=5)
    print('Status:', req.status)
    print('Content-Type:', req.headers.get('Content-Type'))
    print('Content-Disposition:', req.headers.get('Content-Disposition'))
    body = req.read(500)
    print('Body starts with:', body[:100])
except Exception as e:
    print('Error:', e)
