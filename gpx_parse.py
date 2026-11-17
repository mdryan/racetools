#!/usr/bin/env python3

# Standard library
import argparse
import json
import math
import re
from datetime import datetime, timedelta

# Third party
import geopy.distance
import gpxpy
import lxml.etree
import openpyxl
import pandas as pd
import pathvalidate
import plotly
import plotly.express as px
import pwinput
import requests
from openpyxl.utils.dataframe import dataframe_to_rows
from scipy.ndimage import gaussian_filter1d
from scipy.signal import savgol_filter
from scipy.spatial.distance import cdist
from specialsauce.sources.strava import gap_speed_factor

#TODO: close times for checkpoints
#TODO: services for checkpoints
#TODO: gpx extensions to save checkpoint data
#TODO: why doesn't image export work?
#TODO: add a fatigue factor that lets GAP fall during race
#TODO: add a terrain factor for each leg that allows GAP to vary
#TODO: race start time should be the arrive time at START - allow dwell at start to account for difference between gun and race times
#TODO: if START dwell is 0 then don't display either it or start arrival time
#TODO: do not allow dwell at FINISH
#TODO: do not display depart at FINISH
#TODO: allow loading / displaying additional tracks on the same map (but only one as working race)
#TODO: work better with missing elevation data
#TODO: missing elevation on the first trackpoint(s) is filled with 0m (prev_elev = 0), creating a fake climb that inflates D+ and time estimates - interpolate instead
#TODO: replace bare `except:` in input loops with `except ValueError:` - they swallow Ctrl+C (re-prompts) and Ctrl+D (infinite busy loop)
#TODO: handle non-numeric input in 'Select waypoint to delete' and 'Select track/route number' - int(input()) currently crashes, losing unsaved work
#TODO: duplicate consecutive trackpoints give delta_dist 0 -> slope NaN -> NaN passing time -> timedelta(seconds=nan) raises ValueError in estimate_times()

# See https://www.gpxz.io/blog/gpx-file-to-pandas for adding missing elev from DEM model
# See https://github.com/changtimwu/gpx-checkpoint-inserter (as well as gpxpy itself) for algorithm to simplify track log
# See https://logiqx.github.io/gps-wizard/gpx/extensions.html for existing known gpx extensions
# See https://pypi.org/project/garminconnect/ for how to get data from Garmin Connect
# See https://github.com/mshroyer/coursepointer/blob/main/docs/point_types.md for Garmin supported waypoint types

GPX_TOOLS_NS = 'http://mdryan.net/schema/gpx_tools/1.1'
GPX_TOOLS_TAG = f'{{{GPX_TOOLS_NS}}}gpx_tools'

# Experimental: interact with the Ultra Field races database (https://www.vertikalpuntamartin.it/ultraField).
# Disabled by default; set to True to show the UDB option in the main menu.
ENABLE_ULTRAFIELD_DB = False

# Not used yet: lookup table for the planned "allow specifying type of waypoint" feature (see todo in write_gpx())
waypoint_types = {
    'generic': {'name': 'Generic', 'description': 'Generic', 'icon': '📍'},
    'summit': {'name': 'Summit', 'description': 'Summit', 'icon': '⛰️'},
    'valley': {'name': 'Valley', 'description': 'Valley', 'icon': '🏞️'},
    'water': {'name': 'Water', 'description': 'Water', 'icon': '💧'},
    'food': {'name': 'Food', 'description': 'Food', 'icon': '🥣'},
    'danger': {'name': 'Danger', 'description': 'Danger', 'icon': '⚠️'},
    'first_aid': {'name': 'FirstAid', 'description': 'First Aid', 'icon': '🏥'},
    'fourth_category': {'name': 'FourthCategory', 'description': 'Fourth Category', 'icon': '④'},
    'third_category': {'name': 'ThirdCategory', 'description': 'Third Category', 'icon': '③'},
    'second_category': {'name': 'SecondCategory', 'description': 'Second Category', 'icon': '②'},
    'first_category': {'name': 'FirstCategory', 'description': 'First Category', 'icon': '①'},
    'hors_category': {'name': 'HorsCategory', 'description': 'Hors Category', 'icon': '⏲'},
    'sprint': {'name': 'Sprint', 'description': 'Sprint', 'icon': '🏃'},
    'campsite': {'name': 'Campsite', 'description': 'Campsite', 'icon': '🏕️'},
    'aid_station': {'name': 'AidStation', 'description': 'Aid Station', 'icon': '🏁'},
    'rest_area': {'name': 'RestArea', 'description': 'Rest Area', 'icon': '🛌'},
    'general_distance': {'name': 'GeneralDistance', 'description': 'General Distance', 'icon': 'KM'},
    'service': {'name': 'Service', 'description': 'Service', 'icon': '🔧'},
    'energy_gel': {'name': 'EnergyGel', 'description': 'Energy Gel', 'icon': '🍫'},
    'sports_drink': {'name': 'SportsDrink', 'description': 'Sports Drink', 'icon': '🍼'},
    'mile_marker': {'name': 'MileMarker', 'description': 'Mile Marker', 'icon': 'MP'},
    'checkpoint': {'name': 'Checkpoint', 'description': 'Checkpoint', 'icon': 'CP'},
    'shelter': {'name': 'Shelter', 'description': 'Shelter', 'icon': '⛺'},
    'meeting_spot': {'name': 'MeetingSpot', 'description': 'Meeting Spot', 'icon': '📍'},
    'overlook': {'name': 'Overlook', 'description': 'Overlook', 'icon': ' 📷'},
    'toilet': {'name': 'Toilet', 'description': 'Toilet', 'icon': '🚻'},
    'shower': {'name': 'Shower', 'description': 'Shower', 'icon': '🚿'},
    'gear': {'name': 'Gear', 'description': 'Gear', 'icon': '👟'},
    'sharp_curve': {'name': 'SharpCurve', 'description': 'Sharp Curve', 'icon': '⮌'},
    'steep_incline': {'name': 'SteepIncline', 'description': 'Steep Incline', 'icon': '🆙'},
    'tunnel': {'name': 'Tunnel', 'description': 'Tunnel', 'icon': '🚇'},
    'bridge': {'name': 'Bridge', 'description': 'Bridge', 'icon': '🌉'},
    'obstacle': {'name': 'Obstacle', 'description': 'Obstacle', 'icon': '🚧'},
    'crossing': {'name': 'Crossing', 'description': 'Crossing', 'icon': '🚦'},
    'store': {'name': 'Store', 'description': 'Store', 'icon': '🏪'},
    'transition': {'name': 'Transition', 'description': 'Transition', 'icon': '🔃'},
    'navaid': {'name': 'Navaid', 'description': 'Navaid', 'icon': '🧭'},
    'transport': {'name': 'Transport', 'description': 'Transport', 'icon': '🚌'},
    'alert': {'name': 'Alert', 'description': 'Alert', 'icon': '🕭'},
    'info': {'name': 'Info', 'description': 'Info', 'icon': '🛈'}
}

track = pd.DataFrame()
waypoints = pd.DataFrame()
race = {
    'name': None,
    'start': None,
    'pace': None,
}

def set_race_name() -> None:
    global race
    while True:
        name = input('Race name: ')
        name = name.strip()
        if name != '':
            break
    race['name'] = name


# TODO: there is no menu option to change the race start time once it is set (from the GPX file or the first prompt) -
#       this is only called when race['start'] is None
def set_race_start() -> None:
    global race
    race['start'] = input_datetime('Race start date and time in iso format (YYYY-MM-DD HH:MM): ')


# TODO: check the pace is in a plausible range (e.g. 2:30 - 20:00 per km) - a plain number is taken as seconds,
#       so typing 6 meaning 6 min/km gives 6 s/km and absurd time estimates without any warning
def set_race_pace() -> None:
    global race
    race['pace'] = input_secs('Target pace on flat ground in minutes and seconds (MM:SS): ')


def input_datetime(prompt = 'Enter iso date and time (YYYY-MM-DD HH:MM:SS): ') -> datetime:
    """
    Args:
        prompt (str, optional): _description_. Defaults to 'Enter iso date and time (YYYY-MM-DD HH:MM:SS): '.

    Returns:
        datetime: _description_
    """
    while True:
        try:
            dt = datetime.fromisoformat(input(prompt))
        except:
            continue
        return dt

def input_optional_datetime(prompt: str, current: 'datetime | None' = None) -> 'datetime | None | str':
    """
    Prompt for an optional datetime.
    - empty input  -> returns '' (meaning "leave unchanged")
    - 'clear'      -> returns None (meaning "remove the value")
    - valid ISO    -> returns a datetime
    Loops until the user enters something parseable, 'clear', or empty.
    """
    while True:
        raw = input(prompt).strip()
        if raw == '':
            return ''
        if raw.lower() == 'clear':
            return None
        try:
            return datetime.fromisoformat(raw)
        except ValueError:
            print(f"  Could not parse {raw!r} as a datetime. "
                  f"Use 'YYYY-MM-DD HH:MM', 'clear' to remove, or blank to keep unchanged.")
            
def input_secs(prompt = 'Enter time as minutes and seconds MM:SS', allow_zero = False, default = '', allow_blank = False) -> 'int | None':
    """Prompt for a duration as MM:SS, HH:MM:SS or a plain number of seconds (quicker for dwell times, e.g. 90).
    With allow_blank=True a blank answer returns None (no value)."""
    while True:
        try:
            ss = input(prompt)
            if ss == '' and allow_blank:
                return None
            hh = 0
            mm = 0
            if ss == '':
                ss = default
            if ':' in ss:
                if(ss.count(':') == 1):
                    mm, ss = ss.split(':', 2)
                else:
                    hh, mm, ss = ss.split(':', 3)
                    if int(ss) >= 60:
                        raise ValueError('Seconds must be less than 60')
                    if int(mm) >= 60:
                        raise ValueError('Minutes must be less than 60')
                    if int(hh) >= 24:
                        raise ValueError('Hours must be less than 24')
            secs = int(hh) * 60 * 60 + int(mm) * 60 + int(ss)
        except(ValueError):
            continue
        if allow_zero:
            if 0 <= secs < 24*60*60:
                return secs
        else:
            if 0 < secs < 24*60*60:
                return secs


def input_positive_int(prompt = 'Enter an integer: ', allow_zero = False, default = '') -> int:
    while True:
        try:
            i = input(prompt)
            if i == '':
                i = default
            i = int(i)
        except(ValueError):
            continue
        if allow_zero:
            if i >= 0:
                return i
        else:
            if i > 0:
                return i


def input_int(prompt = 'Enter an integer: ', default = '') -> int:
    while True:
        try:
            i = input(prompt)
            if i == '':
                i = default
            i = int(i)
        except(ValueError):
            continue
        return i


def format_time(time_s: float) -> str:
    try:
        mins, secs = divmod(int(time_s), 60)
        hrs, mins = divmod(mins, 60)
        if hrs > 0:
            return f"{hrs:02d}:{mins:02d}:{secs:02d}"
        else:
            return f"{mins:02d}:{secs:02d}"
    except:
        return ''


def excel_sheet_title(name: str) -> str:
    """Make a race name valid as an Excel sheet title."""
    title = name.replace('[', '(').replace(']', ')')   # keep the meaning of "[100K]"
    title = re.sub(r'[*?:/\\]', '-', title)             # other characters Excel forbids
    title = title.strip(" '")[:31].rstrip(" '")         # no edge apostrophes, 31-char limit
    return title or 'Race'                              # never empty


def print_gpx_info(gpx_path: str, gpx: gpxpy.gpx.GPX) -> None:
    print(f"Filename: {gpx_path}")
    print(f"Waypoints: {len(gpx.waypoints)}")
    print(f"Tracks: {len(gpx.tracks)}")
    for t in gpx.tracks:
        print(f"   {t.name}- {len(t.segments)}")
        for s in t.segments:
            print(f"      {len(s.points)} points")
    print(f"Routes: {len(gpx.routes)}")
    for r in gpx.routes:
        print(f"   {r.name}: {len(r.points)}")


def load_track(gpx_track: gpxpy.gpx.GPXTrack) -> None:
    global track
    points = []
    missing_elev = 0
    prev_elev = 0
    for segment in gpx_track.segments:
        for p in segment.points:
            if p.elevation == None:
                missing_elev += 1
                p.elevation = prev_elev
#                print(F"Missing elevation data: <trkpt lat=\"{p.latitude}\" lon=\"{p.longitude}\">")
            #todo: insert as N/A and then interpolate later instead of just using previous elevation (which will make it look like there's a flat section in the route)
            points.append({
                'lon': p.longitude,
                'lat': p.latitude,
                'elev': p.elevation,
                'time': p.time
            })
            prev_elev = p.elevation
    track = pd.DataFrame.from_records(points)
    if missing_elev > 0:
        print(F"*** Warning {missing_elev} points are missing elevation data ***")
    process_track()


def load_route(gpx_route: gpxpy.gpx.GPXRoute) -> None:
    global track
    points = []
    missing_elev = 0
    prev_elev = 0
    for p in gpx_route.points:
        if p.elevation == None:
            missing_elev += 1
            p.elevation = prev_elev
#            print(F"Missing elevation data: <trkpt lat=\"{p.latitude}\" lon=\"{p.longitude}\">")
            #todo: insert as N/A and then interpolate later instead of just using previous elevation (which will make it look like there's a flat section in the route)
        points.append({
            'lon': p.longitude,
            'lat': p.latitude,
            'elev': p.elevation,
        })
        prev_elev = p.elevation
    track = pd.DataFrame.from_records(points)
    if missing_elev > 0:
        print(F"*** Warning {missing_elev} points are missing elevation data ***")
    process_track()


def process_track() -> None:
    global track
    # Calculate euclidean distances between individual points and total distance for each point
    coords = [(p.lat, p.lon, p.elev) for p in track.itertuples()]
    track['delta_dist'] = [0] + [math.sqrt(geopy.distance.distance(from_[:2], to[:2]).m**2 + (from_[2] - to[2])**2) for from_, to in zip(coords[:-1], coords[1:])]
    track['dist'] = track.delta_dist.cumsum()

    # Calculate elevation differences between points
    elevs = [(p.elev) for p in track.itertuples()]
    track['delta_elev'] = [0] + [to - from_ for from_, to in zip(elevs[:-1], elevs[1:])]
    track['elev_gain'] = track[track['delta_elev'] > 0]['delta_elev'].cumsum()
    track['elev_loss'] = track[track['delta_elev'] < 0]['delta_elev'].cumsum()
    track['elev_gain'] = track['elev_gain'].ffill()
    track['elev_loss'] = track['elev_loss'].ffill()
    track['elev_gain'] = track['elev_gain'].fillna(0)
    track['elev_loss'] = track['elev_loss'].fillna(0)
    track['slope'] = track['delta_elev'] / track['delta_dist']

    # Re-calculate estimated passing times
    estimate_times()


def load_waypoints(gpx_waypoints: list[gpxpy.gpx.GPXWaypoint]):
    global waypoints
    points = []
    # TODO: Add beginning and end of each segment instead of whole track
    # TODO: Warn (or do not add altitude) if waypoint altitude is zero
    for wp in gpx_waypoints:
        point = {
            'name': str(wp.name).strip(),
            'lon': wp.longitude,
            'lat': wp.latitude,
            'elev': wp.elevation,
            'dwell': None,
            'close': None
        }
        for extension in wp.extensions:
            if extension.tag == GPX_TOOLS_TAG:
                dwell_txt = extension.attrib.get('dwell', None)
                if dwell_txt and ':' in dwell_txt:
                    if(dwell_txt.count(':') == 1):
                        hh = 0
                        mm, ss = dwell_txt.split(':', 2)
                    else:
                        hh, mm, ss = dwell_txt.split(':', 3)
                    #todo: make robust against badly formed dwell times (e.g. missing minutes, non-numeric, etc.)
                    point['dwell'] = int(hh) * 60 * 60 + int(mm) * 60 + int(ss)
                point['close'] = pd.Timestamp(gpxpy.gpxfield.parse_time(extension.attrib.get('close', None)))
        points.append(point)
    waypoints = pd.DataFrame.from_records(points)

    if ('dwell' in waypoints.columns):
        waypoints['dwell_txt'] = waypoints.apply(lambda wp: format_time(wp['dwell']), axis=1)
    else:
        waypoints['dwell_txt'] = ''

    if len(waypoints) > 0:
        match_waypoints_to_track()

    if len(waypoints) == 0 or geopy.distance.distance((waypoints.at[0, 'lat'], waypoints.at[0, 'lon']), (track.at[0, 'lat'], track.at[0, 'lon'])) > 0.05: # 50m
        print("*** Warning: No START waypoint found, adding one at beginning of track ***")
        waypoints = pd.concat([waypoints, pd.DataFrame([{
            'name': 'START',
            'lon': track.iloc[0].lon,
            'lat': track.iloc[0].lat,
            'elev': track.iloc[0].elev,
            'dwell': None,
            'dwell_txt': '',
            'close': pd.NaT
        }])])
        waypoints.reset_index(inplace=True, drop=True)
        match_waypoints_to_track()
    if geopy.distance.distance((waypoints.at[len(waypoints)-1, 'lat'], waypoints.at[len(waypoints)-1, 'lon']), (track.at[len(track)-1, 'lat'], track.at[len(track)-1, 'lon'])) > 0.05: # 50m
        print("*** Warning: No END waypoint found, adding one at end of track ***")
        waypoints = pd.concat([waypoints, pd.DataFrame([{
            'name': 'END',
            'lon': track.iloc[-1].lon,
            'lat': track.iloc[-1].lat,
            'elev': track.iloc[-1].elev,
            'dwell': None,
            'dwell_txt': '',
            'close': pd.NaT
        }])])
        waypoints.reset_index(inplace=True, drop=True)
        match_waypoints_to_track()

def add_waypoint() -> None:
    global waypoints
    new_wp_name = input('New waypoint name: ')
    while True:
        try:
            new_wp_distance = float(input('New waypoint distance (km): ')) * 1000
            if 0 < new_wp_distance < track.iloc[len(track)-1].dist:
                break
        except:
            continue
    # TODO: Optionally close time, services, dwell...
    closest_trackpoint = track.iloc[(track['dist'] - new_wp_distance).abs().argsort()[:1]].copy()
    closest_trackpoint['name'] = new_wp_name
    closest_trackpoint['dwell'] = None
    closest_trackpoint['dwell_txt'] = ''
    closest_trackpoint['close'] = pd.NaT
    waypoints = pd.concat([waypoints, closest_trackpoint[['lon', 'lat', 'elev', 'name', 'dist', 'dwell', 'dwell_txt', 'close']]])
    waypoints.reset_index(inplace=True, drop=True)
    match_waypoints_to_track()


def edit_waypoint(row) -> None:
    global waypoints

    # Name
    new_name = input(F"New waypoint name [{waypoints.at[row, 'name']}]: ")
    if new_name != '':
        waypoints.at[row, 'name'] = new_name

    # TODO: allow changing distance

    # Elevation
    while True:
        try:
            new_elev = input(F"New elevation [{waypoints.at[row, 'elev']}]: ")
            if new_elev == '':
                break
            if -432 <= float(new_elev) <= 8848.89: # Shore of Dead Sea to summit of Everest would be a tough race! 
                waypoints.at[row, 'elev'] = float(new_elev)
                break
        except ValueError:
            continue

    # Dwell time
    is_finish = waypoints.at[row, 'dist'] == track.iloc[-1].dist
    new_dwell = prompt_dwell_time(waypoints.at[row, 'dwell'], is_finish)
    waypoints.at[row, 'dwell'] = new_dwell
    waypoints.at[row, 'dwell_txt'] = '' if new_dwell is None else format_time(new_dwell)

    # Close time
    edit_close_time(row)

    # Re-match waypoints to track in case elevation or distance changed
    match_waypoints_to_track()


def prompt_close_time(current) -> 'pd.Timestamp | None | str':
    """
    Prompt for a close time, given the current value.
    Returns:
        ''      -> unchanged
        None    -> clear
        Timestamp -> new value
    """
    shown = current.isoformat(sep=' ', timespec='minutes') if isinstance(current, pd.Timestamp) else 'none'
    raw = input_optional_datetime(
        f"Close time [{shown}] (blank to keep, 'clear' to remove): ",
        current=current,
    )
    if raw == '' or raw is None:
        return raw
    return pd.Timestamp(raw)

def edit_close_time(row) -> None:
    global waypoints
    result = prompt_close_time(waypoints.at[row, 'close'])
    if result == '':
        return
    waypoints.at[row, 'close'] = pd.NaT if result is None else result   # None clears, Timestamp sets

def edit_close_times() -> None:
    global waypoints
    for i in waypoints.index:
        print(f"\nWaypoint {i}: {waypoints.at[i, 'name']}")
        edit_close_time(i)

def drop_waypoint(row) -> None:
    global waypoints
    waypoints.drop(index=row, inplace=True)
    waypoints.reset_index(inplace=True, drop=True)
    match_waypoints_to_track()


def match_waypoints_to_track() -> None:
    global waypoints, track

    # remove existing waypoint references from track data
    track['WP'] = pd.NA

    # We don't allow multiple waypoints to match a single trackpoint so keep a list of those available and remove once matched
    # Assuming waypoints are in distance order this tends to place them in the right place along the track where track overlaps
    # with itself (e.g. at start/finish of a loop) since the closest trackpoint to the first waypoint will usually be on the first loop and then that trackpoint will be removed from the list of available trackpoints for the second waypoint which will then match to the second loop
    unused_trackpoints = track.index.tolist()

    for i in waypoints.index:
        distance_to_trackpoints = cdist(
            waypoints[['lon', 'lat']].filter(items=[i], axis=0),
            track[['lon', 'lat']].filter(items=unused_trackpoints, axis=0),
            metric="euclidean"
        )
        pos = int(distance_to_trackpoints.argmin(axis=1)[0])   # position within the filtered subset
        track_idx = unused_trackpoints[pos]                    # <-- the actual track index label

        distance = geopy.distance.distance(
            (waypoints.at[i, 'lat'], waypoints.at[i, 'lon']),
            (track.at[track_idx, 'lat'], track.at[track_idx, 'lon'])
        )
        if distance < 0.55:  # 550m
            waypoints.at[i, 'dist']      = track.at[track_idx, 'dist']
            waypoints.at[i, 'elev_gain'] = track.at[track_idx, 'elev_gain']
            waypoints.at[i, 'elev_loss'] = track.at[track_idx, 'elev_loss']
            waypoints.at[i, 'trkpt_elev']= track.at[track_idx, 'elev']
            track.at[track_idx, 'WP']    = i
            unused_trackpoints.pop(pos)
        else:
            print(F"Warning waypoint {waypoints.at[i, 'name']} is {distance.m:.0f}m from nearest trackpoint")

    # If waypoint didn't have an elevation then use elevation of closest trackpoint
    waypoints['elev'] = waypoints['elev'].fillna(waypoints['trkpt_elev'])

    # Sort waypoints by distance along track
    waypoints.sort_values(by=['dist'], ignore_index=True, inplace=True)
    
    #calculate distances between waypoints
    dists = [(p.dist) for p in waypoints.itertuples()]
    waypoints['leg_dist'] = [0] + [to - from_ for from_, to in zip(dists[:-1], dists[1:])]
    elev_gains = [(p.elev_gain) for p in waypoints.itertuples()]
    waypoints['leg_elev_gain'] = [0] + [to - from_ for from_, to in zip(elev_gains[:-1], elev_gains[1:])]
    elev_losses = [(p.elev_loss) for p in waypoints.itertuples()]
    waypoints['leg_elev_loss'] = [0] + [to - from_ for from_, to in zip(elev_losses[:-1], elev_losses[1:])]

    #reset time estimates since waypoints have changed
    estimate_times()

def prompt_dwell_time(current_secs, is_finish=False) -> 'int | None':
    if is_finish:
        return 0
    if pd.notna(current_secs):
        shown = format_time(current_secs)
        return input_secs(f"Dwell [{shown}] (blank to keep): ", allow_zero=True, default=shown)
    return input_secs("Dwell [none] (MM:SS, blank to leave unset): ", allow_zero=True, allow_blank=True)


def set_dwell_times(reset = False) -> None:
    global waypoints
    last_dist = track.iloc[len(track)-1].dist
    for i, wp in waypoints.iterrows():
        if wp['dist'] == last_dist:
            # Dwell time must always be zero for finish line waypoint
            dwell = 0
        elif reset:
            if pd.notna(wp['dwell']):
                shown = format_time(wp['dwell'])
                dwell = input_secs(f"Enter dwell time for {wp['name']} in minutes and seconds [{shown}] (blank to keep): ", allow_zero=True, default=shown)
            else:
                dwell = input_secs(f"Enter dwell time for {wp['name']} in minutes and seconds [none] (blank to leave unset): ", allow_zero=True, allow_blank=True)
        elif pd.isna(wp['dwell']):
            dwell = input_secs(
                f"Enter dwell time for {wp['name']} in minutes and seconds (MM:SS): ",
                allow_zero=True,
            )
        else:
            dwell = wp['dwell']
        waypoints.at[i, 'dwell'] = dwell
    waypoints['dwell_txt'] = waypoints.apply(lambda wp: '' if pd.isna(wp['dwell']) else format_time(wp['dwell']), axis=1)


# TODO: allow multiple tracks
def plot_map(map_style: str='open-street-map') -> plotly.graph_objs._figure.Figure:
    fig = px.line_map(
        track,
        lat='lat',
        lon='lon',
        hover_data={'elev': ':,.0f', 'dist': ':,.0f', 'lat': False, 'lon': False},
        map_style=map_style,
        zoom=12
    )
    fig.update_layout(margin={"r":0,"t":0,"l":0,"b":0})
    fig.update_traces(line=dict(width=5))
    fig.add_scattermap(
        lat=waypoints['lat'],
        lon=waypoints['lon'],
        text=waypoints['name'],
        marker=dict(size=16),
        textfont=dict(size=16, color='white', weight='bold'),
        textposition='top right',
        mode='markers+text',
        hoverinfo='none',
        showlegend=False
    )
    return fig


def plot_elevation(ratio=10, min_elev=None, max_elev=None) -> plotly.graph_objs._figure.Figure:
    if min_elev == None:
        min_elev = track['elev'].min() * 0.9
    if max_elev == None:
        max_elev = track['elev'].max() * 1.1 
    fig = px.area(track, x='dist', y='elev', template='plotly', width=1800)
    fig.update_yaxes(
        title='',
        ticksuffix='m',
        tickfont=dict(size=24),
        scaleanchor='x',
        scaleratio=ratio,
        range=[min_elev, max_elev],
        constrain="domain",
        constraintoward="bottom",
    )
    fig.update_xaxes(
        title='',
        showgrid=False,
        minallowed=0,
        maxallowed=track.iloc[len(track)-1].dist,
    )
    labels = waypoints.dropna(subset=['dist'])
    if len(labels) > 0:
        fig.update_xaxes(
            tickmode='array',
            tickvals=labels['dist'],
            ticktext=round(labels['dist']/1000, 1),
            tickfont=dict(size=24),
            showgrid=True,
            gridcolor='salmon',
            gridwidth=3,
            griddash='dot'
        )
        fig.add_scatter(
            x=labels['dist'],
            y=labels['elev'],
            text=round(labels['elev']),
            marker=dict(size=20),
            textfont=dict(size=24),
            textposition='top left',
            hoverinfo='text',
            mode='markers+text',
            hovertext=labels['name'],
            showlegend=False,
            xaxis='x2'
        )
        fig.update_layout(xaxis2=dict(
            tickmode='array',
            tickvals=labels['dist'],
            ticktext=labels['name'],
            tickfont=dict(size=24),
            overlaying='x',
            side='top',
            anchor='y',
            minallowed=0,
            maxallowed=track.iloc[len(track)-1].dist,
            showgrid=False
        ))
    return fig


def estimate_times() -> None:
    global waypoints

    if race['pace'] == None:
        waypoints['depart'] = None
        waypoints['arrive'] = None
        waypoints['avg_pace'] = None
        return

    if race['start'] == None:
        set_race_start()

    set_dwell_times()

    cumulative_dwell_times = waypoints['dwell'].cumsum()
    times = (track['delta_dist'] * race['pace'] / 1000  * gap_speed_factor(track['slope'])).cumsum()
    times[0] = 0
    trkpts = track.dropna(subset=['WP']).index.tolist()
    if len(trkpts) != len(waypoints):
        print("*** Warning number of waypoints does not match number of trackpoints with waypoints ***")
        waypoints['depart'] = None
        waypoints['arrive'] = None
        waypoints['avg_pace'] = None
        return

    wp_times = pd.Series(times.loc[trkpts].to_numpy(), index=waypoints.index)
    wp_depart = wp_times + cumulative_dwell_times
    wp_arrive = wp_times + cumulative_dwell_times - waypoints['dwell']
    waypoints['depart'] = [race['start'] + timedelta(seconds=depart) for depart in wp_depart]
    waypoints['arrive'] = [race['start'] + timedelta(seconds=arrive) for arrive in wp_arrive]
    waypoints['depart'] = waypoints['depart'].dt.round('s')
    waypoints['arrive'] = waypoints['arrive'].dt.round('s')

    # Calculate average pace for each leg
    leg_times = [0] + [to - from_ for from_, to in zip(wp_arrive[:-1], wp_arrive[1:])]
    avg_pace = leg_times / waypoints['leg_dist'] * 1000
    waypoints['avg_pace'] = [format_time(a) for a in avg_pace]

    # If waypoint at finish line then do not show depart or dwell time
    if geopy.distance.distance((waypoints.at[len(waypoints)-1, 'lat'], waypoints.at[len(waypoints)-1, 'lon']), (track.at[len(track)-1, 'lat'], track.at[len(track)-1, 'lon'])) < 0.05: # 50m
        waypoints.at[len(waypoints)-1, 'depart'] = None
        waypoints.at[len(waypoints)-1, 'dwell_txt'] = ''

    # If waypoint at start line and no dwell time then do not show arrive or dwell time
    # (dwell time at start line is usually used to account for difference between gun time and chip time, but if it's zero then it's cleaner to not show it at all)
    if waypoints.iloc[0]['dist'] == 0 and waypoints.iloc[0]['dwell'] == 0:
        waypoints.at[0, 'arrive'] = None
        waypoints.at[0, 'dwell_txt'] = ''


def write_excel() -> None:
    if race['name'] == None:
        set_race_name()
    basename = pathvalidate.sanitize_filename(race['name'])
    today = datetime.today().strftime('%Y%m%d%H%M')

    # Create workbook and worksheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = excel_sheet_title(race['name'])

    export = waypoints[['dist', 'name', 'leg_dist', 'elev', 'leg_elev_gain', 'leg_elev_loss', 'avg_pace', 'arrive', 'dwell', 'close']]
    export['dist'] = round(waypoints['dist']/1000, 1)
    export['leg_dist'] = round(waypoints['leg_dist']/1000, 1)
    export['dwell'] = waypoints['dwell'] / (24*60*60)
    
    ws.append(['Km', 'Name', 'Dist', 'Elev', 'D+', 'D-', 'Pace', 'ETA', 'Dwell', 'Close'])
    for row in dataframe_to_rows(export, index=False, header=False):
        ws.append(row)

    # Define styles
    header_font = openpyxl.styles.Font(bold=True, size=11, name='Calibri')
    header_fill = openpyxl.styles.PatternFill(start_color="b2b2b2", end_color="366092", fill_type="solid")
    odd_fill = openpyxl.styles.PatternFill(start_color="dddddd", end_color="dddddd", fill_type="solid")
    data_font = openpyxl.styles.Font(size=11, name='Calibri')
    thin = openpyxl.styles.Side(style='thin')
    border = openpyxl.styles.Border(left=thin, right=thin, top=thin, bottom=thin)
    center_align = openpyxl.styles.Alignment(horizontal='center')

    # Header styles
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        if cell.column != 2: # Everything but name column
            cell.alignment = center_align

    # Row striping
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        if row[0].row % 2 == 1:
            for cell in row:
                cell.fill = odd_fill

    # Data styles
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = data_font
            cell.border = border
            if cell.column != 2: # Everything but name column
                cell.alignment = center_align
            if cell.column in [1, 3]: # Distance
                cell.number_format = '0.0'
            if cell.column in [4, 6]: # Elevation / elevation loss
                cell.number_format = '0'
            if cell.column in [5]: # Elevation gain
                cell.number_format = '+0'
            if cell.column in [8, 10]: # Arrive and close times
                cell.number_format = 'hh:mm'
            if cell.column in [9]: # Dwell time
                cell.number_format = '[mm]:ss'

    # Column widths
    for column in ws.columns:
        if column[0].column_letter in ['A','C']: # Distance columns
            ws.column_dimensions[column[0].column_letter].width = 5
        if column[0].column_letter == 'B': # Name column
            ws.column_dimensions[column[0].column_letter].width = 20
        if column[0].column_letter in ['D','E','F', 'G','H','I','J']: # Elevation and time columns
            ws.column_dimensions[column[0].column_letter].width = 6

    # Don't show anything for delta columns if WP at 0km
    if ws.cell(row=2, column=1).value == 0:
        for col in [3, 5, 6]:
            ws.cell(row=2, column=col).value = ''

    # Actually save the file
    wb.save(F"{basename}.{today}.xlsx")

def write_gpx() -> None:
    if race['name'] == None:
        set_race_name()
    basename = pathvalidate.sanitize_filename(race['name'])
    today = datetime.today().strftime('%Y%m%d%H%M')
    # Write GPX file with track and waypoints
    gpx = gpxpy.gpx.GPX()
    gpx.creator = "Matt's GPX tools"
    gpx.name = race['name']
    gpx.nsmap = {'mdryan.net': GPX_TOOLS_NS}
    if race['start']:
        gpx.metadata_extensions.append(lxml.etree.Element(GPX_TOOLS_TAG, attrib={'start': race['start'].isoformat()}))

    gpx_track = gpxpy.gpx.GPXTrack()
    gpx_segment = gpxpy.gpx.GPXTrackSegment()
    gpx.tracks.append(gpx_track)
    gpx_track.segments.append(gpx_segment)
    for i, trkpt in track.iterrows():
        gpx_segment.points.append(gpxpy.gpx.GPXTrackPoint(
            trkpt.lat,
            trkpt.lon,
            elevation=round(trkpt.elev, 2)
        ))
    for i, wp in waypoints.iterrows():
        #todo: UltraField food and services metadata in extension
        #todo: Allow specifying type of waypoint
        attribs = {}
        if isinstance(wp['dwell_txt'], str) and wp['dwell_txt'] != '':
            attribs['dwell'] = wp['dwell_txt']
        if pd.notna(wp['close']):
            attribs['close'] = wp['close'].isoformat()
        waypoint = gpxpy.gpx.GPXWaypoint(
            wp['lat'],
            wp['lon'],
            elevation=round(wp['elev'], 2),
            name=wp['name'],
            type='AID STATION',
        )
        if len(attribs) > 0:
            waypoint.extensions.append(lxml.etree.Element(GPX_TOOLS_TAG, attrib=attribs))
        gpx.waypoints.append(waypoint)
    
    with open(F"{basename}.{today}.gpx", "w", encoding='utf-8') as file:
        file.write(gpx.to_xml())

# TODO: round() raises on NaN dist / elevation for a waypoint that didn't match the track - skip or warn instead
# TODO: escape or reject ';' and '|' in race and waypoint names - they are the field and checkpoint separators
def ultra_string() -> str:
    if race['name'] == None:
        set_race_name()
    cps = []
    for i, wp in waypoints.iterrows():
        if wp['dist']==0:
            continue
        if isinstance(wp['close'], pd.Timestamp):
            close = 't!' + format_time(wp['close'].hour *60 + wp['close'].minute)
        else:
            close = ''
        cps.append(';'.join([
            str(round(wp['dist']/10)),
            wp['name'],
            str(round(wp['elev'])),
            str(round(wp['leg_elev_gain'])),
            str(round(wp['leg_elev_loss'] * -1)),
            close
        ]))
    return F"V02;{race['name']};;|{'|'.join(cps)}"


def ultra_db() -> None:
    session = requests.Session()
    # session.get('https://www.vertikalpuntamartin.it/ultraField/')
    # username = input('Username for https://www.vertikalpuntamartin.it/ultraField: ')
    # password = pwinput.pwinput('Password for https://www.vertikalpuntamartin.it/ultraField: ', mask='*')
    # login_url = 'https://www.vertikalpuntamartin.it/UFphpRest/login.php'
    # data = {'user': username, 'password': password}
    # response = session.post(login_url, json=data)
    
    # if response.status_code == requests.codes.ok:
    #     print('Login request was successful')
    #     user = response.json()
    #     print(F"Logged in as {user['displayName']} ({user['id']})")
    # else:
    #     print('Login request failed with status code:', response.status_code)
    #     return

    url = 'https://www.vertikalpuntamartin.it/UFphpRest/paginate.php'    
    params = {
        'page': 1,
        'itemsPerPage': 150,
        'mustSort': 'false',
        'multiSort': 'false',
        'filters': '{}' # possible filter to search for race: {"name like":"%emei%"}'
    }
    response = session.get(url, params=params)
#    print(response.json())
    races = pd.DataFrame.from_records(response.json()['records'], exclude=['user_id'], index=['id'])
    print(races[['race_date', 'name', 'location', 'country', 'created']].to_markdown(headers=['Date', 'Name', 'Location', 'Country', 'Created']))
    while True:
        raceID = input(('Select race ID: '))
        if raceID in races.index:
            break

    url = 'https://www.vertikalpuntamartin.it/UFphpRest/get.php'
    params = {'id': raceID}
    response = session.get(url, params=params)
    cps = json.loads(response.json()['check_points'])
    waypoints = pd.DataFrame.from_records(cps, columns=['distance', 'label', 'ele', 'dplus', 'dminus', 'barrier', 'foods', 'services'])
    print(waypoints)                


def smooth_elevation(mode = "gaussian") -> None:
    global track
    match mode:
        case "gaussian":
            track['elev'] = gaussian_filter1d(track['elev'], sigma=3)
        case "savitzky-golay":
            # The Savitzky-Golay filter is particularly effective for preserving the shape of peaks in your data while reducing noise.
            # It fits a polynomial to a moving window of data points and uses that to estimate the smoothed value.
            # The window_length must be odd and the poly_order must be less than the window length. 
            window_length = 15  # Must be odd
            poly_order = 3  # Polynomial order
            track['elev'] = savgol_filter(track['elev'], window_length, poly_order)
        case _:
            print(F"Unknown smoothing mode {mode}")
            return
    process_track()


def main_menu() -> bool:
    print()
    if len(waypoints) > 0:
        print(waypoints[['dist', 'name', 'elev', 'trkpt_elev', 'leg_dist', 'leg_elev_gain', 'leg_elev_loss', 'avg_pace', 'arrive', 'dwell_txt', 'depart', 'close']].to_markdown(floatfmt=",.0f", headers=['Distance', 'Name', 'Elev', 'Trk El', 'Dist', 'D+', 'D-', 'Pace', 'Arrive', 'Dwell', 'Depart', 'Close']))
    print(f"Track length: {round(max(track['dist'])/1000, 2)}km")
    print(f"Elevation Gain: {round(sum(track[track['delta_elev'] > 0]['delta_elev']))}m")
    print(f"Elevation Loss: {round(sum(track[track['delta_elev'] < 0]['delta_elev']))}m")
    if waypoints.iloc[-1]['arrive'] != None:
        print(f"{waypoints.iloc[-1]['name']} at {waypoints.iloc[-1]['arrive'] - race['start']}")
    if waypoints.dwell.sum() > 0:
        print(f"Total dwell time: {format_time(waypoints.dwell.sum())}")
    options = []
    if len(track) > 0:
        options.append(('A', 'Add waypoint'))
    if len(waypoints) > 0:
        options.append(('D', 'Delete waypoint'))
        options.append(('E', 'Edit waypoint'))
        options.append(('ED', 'Edit waypoint dwell times'))
        options.append(('EC', 'Edit waypoint close times'))
        options.append(('T', 'Estimate times'))
    if len(track) > 0 or len(waypoints) > 0:
        options.append(('L', 'Show elevation'))
        options.append(('S', 'Smooth elevation data'))
        options.append(('M', 'Show map'))
    if len(track) > 0:
        options.append(('W', 'Write gpx file'))
    if len(waypoints) > 0:
        options.append(('U', 'Show UltraField string'))
        if ENABLE_ULTRAFIELD_DB:
            options.append(('UDB', 'Create race on https://www.vertikalpuntamartin.it/ultraField'))
    options.append(('Q', 'Quit'))
    print()
    for o in options:
        print(F" {o[0]} - {o[1]}")
    while True:
        selection = input('Select operation: ').upper()
        if selection in [o[0] for o in options]:
            break
    match selection:
        case 'A':
            add_waypoint()
        case 'D':
            while True:
                row = int(input('Select waypoint to delete: '))
                if row in waypoints.index:
                    break
            drop_waypoint(row)
        case 'E':
            while True:
                try:
                    row = int(input('Select waypoint to edit: '))
                    if row in waypoints.index:
                        break
                except:
                    continue
            edit_waypoint(row)
        case 'ED':
            set_dwell_times(reset=True)
            estimate_times()
        case 'EC':
            edit_close_times()
        case 'M':
            styles = [
                'basic',
                'carto-darkmatter',
                'carto-darkmatter-nolabels',
                'carto-positron',
                'carto-positron-nolabels',
                'carto-voyager',
                'carto-voyager-nolabels',
                'dark',
                'light',
                'open-street-map',
                'outdoors',
                'satellite',
                'satellite-streets',
                'streets',
                'white-bg'
            ]
            i = 1
            for s in styles:
                print(f"{i} - {s}")
                i+=1
            while True:
                try:
                    style = styles[(int(input('Select map style: '))) - 1]
                    if style:
                        break
                except:
                    continue
            plot_map(style).show()
        case 'L':
            ratio = input_positive_int('Vertical exaggeration ratio (default: 10): ', default=10)
            min_elev = input_int('Minimum elevation on plot (default: auto): ', default=track['elev'].min() * 0.9)
            max_elev = input_int('Maximum elevation on plot (default: auto): ', default=track['elev'].max() * 1.1)
            plot_elevation(ratio, min_elev, max_elev).show()
        case 'S':
            modes = [
                'gaussian',
                'savitzky-golay'
            ]
            for i, mode in enumerate(modes, start=1):
                print(f"{i} - {mode}")
            while True:
                try:
                    mode = modes[(int(input("Select smoothing mode: "))) - 1]
                    if mode:
                        break
                except:
                    continue
            smooth_elevation(mode)
            #TODO: do we need (or even want?) to do this here?
            match_waypoints_to_track()
        case 'W':
            write_gpx()
            write_excel()
            # Write elevation and plan as images
            # todo: FIXME
            # fig = plot_elevation(track, waypoints)
            # fig.write_image(F"{basename}.elevation.{today}.png")
            # fig = plot_map(track, waypoints)
            # fig.write_image(F"{basename}.plan.{today}.png")

        case 'U':
            print(ultra_string())
        case 'UDB' if ENABLE_ULTRAFIELD_DB:
            ultra_db()
        case 'T':
            set_race_pace()
            estimate_times()
        case 'Q':
            # TODO: warn about unsaved changes (anything edited since the last W) and ask to confirm before quitting
            return False
    return True


def main(gpx_path: str) -> None:
    global waypoints, track, race
    with open(gpx_path, encoding='utf-8') as f:
        gpx = gpxpy.parse(f)
    
    print_gpx_info(gpx_path, gpx)
    if gpx.name:
        race['name'] = gpx.name

    for extension in gpx.metadata_extensions:
        if extension.tag == GPX_TOOLS_TAG:
            race['start'] = gpxpy.gpxfield.parse_time(extension.get('start', None))   

    print(f"Race name: {race['name']}")
    print(f"Race start: {race['start']}")
        
    options = []
    for t in gpx.tracks:
        options.append(t)
    for r in gpx.routes:
        options.append(r)
    if (len(options) == 0):
        print("GPX file contains neither tracks nor routes")
        return
    if (len(options) == 1):
        selected = 0
    else:
        n = 1
        for o in options:
            print(f"{n}: {type(o)} - {o.name}")
            n += 1
        while True:
            selected = int(input("Select track/route number: "))
            if 0 < selected < n:
                selected -= 1
                break
    if type(options[selected]) == gpxpy.gpx.GPXTrack:
        # Automatic smoothing on load disabled: it only worked for tracks (GPXRoute has no smooth()), applied a
        # 0.4/0.2/0.4 neighbour average to tracks drawn in planning tools as well as recorded ones, and stacked
        # with the S menu filters. TODO: offer gpxpy's weights (gpxpy.gpx.SMOOTHING_RATIO) as an S menu option instead.
        # options[selected].smooth()
        load_track(options[selected])
    elif type(options[selected]) == gpxpy.gpx.GPXRoute:
        load_route(options[selected])
    else:
        print("Unknown gpx data type selected")
        return
    
    load_waypoints(gpx.waypoints)

    while main_menu():
        pass

    
# TODO: consider parser.add_argument('infile', type=argparse.FileType('r'))
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Extract elevation profile and checkpoint data from a gpx file')
    parser.add_argument('filename')
    args = parser.parse_args()
    return args
    

if __name__ == '__main__':
    args = parse_args()
    main(args.filename)
