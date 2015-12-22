# ghosttrip

This is a quickly hacked-together serial port NMEA GPS Emulator for Windows using MFC. I needed a simple piece of software to walk a route of waypoints at a given speed while writing NMEA compatible GPS data to a selected serial port. For me, it was particularily useful for testing/debugging an external GPS logger. As the experiment is done now, I thought someone might find this code useful as a basis for own projects instead of just dummping it.

Features:
- reads KML files i.e. created with Google Earth
- orders the KML waypoints to get a shorter round trip
- start/stop walking with space bar
- correct current position with arrow keys or num block
- double click waypoint to set it as new destination
- speech support (I had only one screen for the PC and the device so this was useful...)

This repo contains code pieces from:

http://www.grinninglizard.com/tinyxml

https://ukhas.org.uk/code:emulator

http://www.naughter.com/enumser.html

https://github.com/samlbest/traveling-salesman
