#! /usr/bin/env python

from json.encoder import JSONEncoder
from os import kill, name
from re import L
from typing import Dict, ItemsView, List, final
from numpy import clongfloat, e, get_printoptions
from numpy.lib.function_base import append
from pandas.io.stata import ValueLabelTypeMismatch
import xmltodict
import pprint
import json
import pandas as pd
import xlsxwriter


print("Start of Python Script")
print("*" * 100)

filepath = input("Enter the file name: ")

my_xml = open(filepath, "r").read()

pp = pprint.PrettyPrinter(indent=4)
new_data = json.dumps(xmltodict.parse(my_xml))

data = json.loads(new_data)


new_line = [
    anim
    for anim in data["config"]["wga_config"]["prox_acl_custom_categories"][
        "prox_acl_custom_category"
    ]
]

(
    url_with_only_url_name,
    url_with_only_regex_name,
    url_with_both_url_name,
    url_with_both_url_name_url,
    url_with_both_url_name_regex,
) = ({}, {}, {}, {}, {})

(
    df_with_actual_regex,
    df_with_both_url,
    df_with_actual_url,
    df_with_regex_name,
    df_with_url_name,
) = ([], [], [], [], [])

NoneType = type(None)

for line in new_line:

    if (
        type(line["prox_acl_custom_category_servers"]) is dict
        and type(line["prox_acl_custom_category_regex_list"]) is dict
    ):
        url_with_both_url_name_url[line["prox_acl_custom_category_name"]] = line[
            "prox_acl_custom_category_servers"
        ]["prox_acl_custom_category_server"]
        url_with_both_url_name_regex[line["prox_acl_custom_category_name"]] = line[
            "prox_acl_custom_category_regex_list"
        ]["prox_acl_custom_category_regex"]
    elif (
        type(line["prox_acl_custom_category_servers"]) is dict
        and type(line["prox_acl_custom_category_regex_list"]) is NoneType
    ):
        url_with_only_url_name[line["prox_acl_custom_category_name"]] = line[
            "prox_acl_custom_category_servers"
        ]["prox_acl_custom_category_server"]
    elif (
        type(line["prox_acl_custom_category_servers"]) is NoneType
        and type(line["prox_acl_custom_category_regex_list"]) is dict
    ):
        url_with_only_regex_name[line["prox_acl_custom_category_name"]] = line[
            "prox_acl_custom_category_regex_list"
        ]["prox_acl_custom_category_regex"]


for key, value in url_with_only_url_name.items():
    df_with_url_name.append(key)
    df_with_actual_url.append(value)

for key, value in url_with_only_regex_name.items():
    df_with_regex_name.append(key)
    df_with_actual_regex.append(value)


ip_in_line_data = [
    anim
    for anim in data["config"]["wga_config"]["prox_acl_policy_groups"]["prox_acl_group"]
]

policy_name = []

only_ip = [anim for anim in ip_in_line_data if "prox_acl_group_ips" in anim]

policy_map = {}

for line in only_ip:
    if line["prox_acl_group_ips"]:
        policy_map[line["prox_acl_group_id"]] = line["prox_acl_group_ips"][
            "prox_acl_group_ip"
        ]
    else:
        policy_map[line["prox_acl_group_id"]] = "No IP Configured"


copy_to_workbook = xlsxwriter.Workbook("Proxy_Rule.xlsx")
dark_blue_header_format = copy_to_workbook.add_format(
    {
        "bg_color": "#00CC66",
        "font_size": 15,
        "bold": True,
        "border": 1,
    }
)


full_border = copy_to_workbook.add_format({"border": 1})
outsheet = copy_to_workbook.add_worksheet(name="POLICY name with IP")
outsheet.write("A1", "Policy Name", dark_blue_header_format)
outsheet.write("B1", "IP Addresses", dark_blue_header_format)


count = 1

for key, value in policy_map.items():

    outsheet.write(count, 0, key, full_border)
    outsheet.write(count, 1, str(value), full_border)
    count = count + 1

outsheet2 = copy_to_workbook.add_worksheet(name="URL Categories with URL allowed")
outsheet2.write("A1", "URL Category Name", dark_blue_header_format)
outsheet2.write("B1", "URL", dark_blue_header_format)

new_count = 0


count_to_monitor = 1

for item in range(len(df_with_url_name)):

    outsheet2.write(count_to_monitor, 0, df_with_url_name[item], full_border)
    outsheet2.write(count_to_monitor, 1, str(df_with_actual_url[item]), full_border)
    count_to_monitor = count_to_monitor + 1


for item in range(len(df_with_regex_name)):

    outsheet2.write(count_to_monitor, 0, df_with_regex_name[item], full_border)
    outsheet2.write(count_to_monitor, 2, str(df_with_actual_regex[item]), full_border)
    count_to_monitor = count_to_monitor + 1

final_monitor = count_to_monitor


for key, value in url_with_both_url_name_url.items():

    outsheet2.write(count_to_monitor, 0, key, full_border)
    outsheet2.write(count_to_monitor, 1, str(value), full_border)
    count_to_monitor = count_to_monitor + 1

for key, value in url_with_both_url_name_regex.items():

    for key2, value2 in url_with_both_url_name_url.items():
        if key == key2:
            outsheet2.write(final_monitor, 2, str(value), full_border)
            final_monitor = final_monitor + 1
            break


copy_to_workbook.close()

print("*" * 100)
print("End of Python Script")
