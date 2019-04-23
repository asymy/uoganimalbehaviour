# Animal Behaviour Analyser

Tool for tracking bites and licks from pre-recorded videos.
Will output an excel spreadsheet with information on the quanity, frequency and duration of behaviours.
Designed for use in the spinal cord group at the University of Glasgow.

## Requirements

- Python3 (3.7.3 used to test)
- PyVlc
- PyQt5
- openpyxl
- VLC (32bit)

## Instructions for installation

- Download [python3](https://www.python.org/)
- Download 32-bit version of [VLC](https://www.videolan.org/vlc/index.en-GB.html)
- open a terminal window
- copy these lines into terminal window and wait for them to install:

`pip install python-vlc`

`pip install openpyxl`

`pip install PyQt5`

- download this project
- run main.py

## Shortcuts

- Control + S = Save
- B = Start Bite
- L = Start Lick
- [/+ = Increase Speed 

## TODO

- [ ] add markers for indicating where the behaviour occurs
- [ ] add support for mouse interaction with these markers
- [x] fix saving the excel sheet
  - [x] format
  - [x] summary sheet
  - [x] indication that it has been saved/option on place to save
- [ ] perhaps .ini file or similar which will allow users to customise their file saving/opening defualt locations
- [ ] add run time/time elapsed?

## Known Issues
