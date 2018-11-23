#!/usr/bin/env python
# -*-coding:utf-8 -*-

import requests,json,Config
from Commons import globals


class ConfigHttp:
    def __init__(self):
        global host, port, timeout
        host = Config.baseUrl
        self.log = globals.log()
        self.headers = {}
        self.params = {}
        self.data = {}
        self.url = None
        self.files = {}

    def set_url(self, url):
        self.url = host + url

    def set_headers(self, header):
        self.headers = header

    def set_params(self, param):
        self.params = param

    def set_data(self, data):
        self.data = data

    def set_files(self, file):
        self.files = file

    # defined http get method
    def get(self):
        try:
            response = requests.get(self.url, params=self.params, headers=self.headers)
            # response.raise_for_status()
            return response
        except Exception as e:
            self.log(e.message)
            return None

    # defined http post method
    def post(self):
        try:
            response = requests.post(self.url, headers=self.headers, data=self.data)
            # response.raise_for_status()
            return response
        except Exception as e:
            self.log(e.message)
            return None