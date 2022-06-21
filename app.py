#!/usr/bin/env python
# -*- coding: utf-8 -*-

from urllib import request, error
from os import environ
import json


def get_response(url):
    try:
        response = request.urlopen(url)
    except error.URLError as e:
        print('Error:', e.reason)
        return None

    if response and response.getcode() == 200:
        raw_data = response.read()
        data = json.loads(raw_data)
        return data
    else:
        print("Error receiving data", response.getcode())
        return None


# printMockResponse() is used to debug
def print_mock_response(data):
    for collection in data['collections']:
        print(f'collections Name:  {collection["name"]}')

        for category in collection['categories']:
            print(f'\ttype: {category["type"]}')

            for reference in category['references']:
                print(f'\t\tid: {reference["id"]}')
                print(f'\t\tmodele/collection: {reference["modele/collection"]}')
                for color in reference['colors']:
                    print(f'\t\t\tcolor variant: {color}')
                print(f'\t\twwp: {reference["wwp"]}')
                print(f'\t\tnomenclature: {reference["nomenclature"]}')
                print(f'\t\tdetails nomenclature: {reference["details nomenclature"]}')
                print(f'\t\tventes: {reference["ventes"]}')
                print(f'\t\tforecast: {reference["forecast"]}')
                print(f'\t\tvisuel: {reference["visuel"]}')


def main():
    mock_api = environ.get('mock_api') or 'http://127.0.0.1:8090/api/v1/datas'
    data = get_response(mock_api)

    if data is not None:
        print_mock_response(data)


if __name__ == '__main__':
    main()
