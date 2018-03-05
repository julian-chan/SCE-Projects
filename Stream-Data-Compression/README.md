# Stream-Data-Compression
This repository contains the code that I developed as a 2017 summer intern for Southern California Edison's (SCE) Advanced Technology Group in Grid Modernization Planning & Technology (GMP&T).

This folder contains a Python script that compresses AutoCAD binary (.dst) files of Synchrophasor data generated every 3 minutes by the company software. Files are compressed by day and stored back into the company database, cutting down 80% of file storage space.  In addition, it contains a GUI for user-friendly access to the back-end compression script. It currently supports directory lookup, data type selection, and parameter error checking.
