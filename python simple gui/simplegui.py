#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jul 28 01:15:05 2022

@author: johnfrost
"""
import PySimpleGUI as sg # pip install pysimplegui

# ------- GUI Definition ----------- #

layout = [ 
    [sg.Text("Input File:"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xls*"),))],
    [sg.Text("Output Folder:"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
    [sg.Exit(), sg.Button("Convert To CSV")]
    ]

window = sg.Window("Excel to CSV Converter", layout)

while True:
    event, values = window.read()
    print(event, values)
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    if event == "Convert To CSV" :
        sg.popup_error("Not yet implemented")
        
window.close()
