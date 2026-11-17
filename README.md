# GPX tools for trail races

A python tool to take gpx files provided in various formats and quality by race organisers and use them more easily for race preparation. A quick hack by a non-python programmer that whilst immature has grown to the extent that it may be of interest to others.

## Features

- [X] Match waypoints to closest trackpoints
- [X] Estimate stage climb, descent and distance based on trackpoint data
- [X] Add waypoints based on distance along track
- [X] Remove selected unnecessary waypoints
- [X] Show map of route (with waypoints)
- [X] Show elevation profile of route (with waypoints)
- [X] Smooth elevation profiles
- [X] Allow a choice of elevation smoothing algorithms
- [X] Export waypoint information to Excel
- [X] Export clean gpx file for navigation during race
- [X] Generate "Custom race: checkpoints string" setting for the [Connect IQ datafield "Ultra Field"](https://apps.garmin.com/en-US/apps/b61ec459-c221-43fc-8b4b-b9730d642368#0)
- [X] Estimate passing and finishing times based on grade adjusted pace and checkpoint dwell times 
- [ ] Apply a fatigue factor during races such that grade adjusted pace falls over time (or rises for people with negative splits on short races)
- [ ] Allow specifying a terrain factor for each section to allow pace to vary independently of grade and fatigue 
- [ ] Allow reducing resolution / smoothing track in plan dimension
- [ ] Split elevation plots across two or more rows for longer races
- [ ] Use SRTM (or other DEM) data for files without elevation data
- [ ] Store additional checkpoint data with waypoints in .gpx file as gpx extensions
- [ ] Show other gpx routes on the map (e.g. other distances, or previous years races)
- [ ] Interact with the [Ultra Field races database](https://www.vertikalpuntamartin.it/ultraField) *(partially implemented: read-only, disabled by default)*
- [ ] Interact with Garmin Connect to allow direct upload of the course
- [ ] Convert to a flask app and host it somewhere
- [ ] Built a totally seperate script to do post match analysis (ideally from the .fit files containing doppler speed estimates)
  - [ ] Compare predicted with actual times (e.g. to estimate impact of fatigue)
  - [ ] Interact with Garmin Connect to avoid manually exporting data
  - [ ] Interact with itra website to plot existing itra score of each finisher by finish time in attempt to predict race score
  - [X] Interact with Weibu (Running π) or other platforms to get checkpoint passing times and rankings

## Installation and getting started

These tools were developed and tested with Python version 3.13.7 on Ubuntu. They should work with minor or no adjustments on similar generations of python on any platform. Regardless of platform users need to be comfortable working with Python and the command line.

### Download the project

```
git clone https://github.com/mdryan/racetools.git
```

(alternatively download and extract the .zip file from https://github.com/mdryan/racetools/archive/refs/heads/main.zip if for some reason you don't have access to git)

### Create a python Virtual Environment in the project directory:

```
cd racetools
python3 -m venv venv
```

### Install dependencies:

```
./venv/bin/pip3 install -r requirements.txt
```

### Run the program

Activate the virtual environment, then run the script:

```
source venv/bin/activate
./gpx_parse.py /path/to/file.gpx
```

(alternatively, without activating the virtual environment, run it with the python interpreter in the virtual environment: `./venv/bin/python3 gpx_parse.py /path/to/file.gpx`)

## Known bugs

- Probably doesn't work when a race course consists of multiple tracks or segments

Plus lots of unknown bugs