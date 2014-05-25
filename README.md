matlab_saveodp
===============

Create Open Document Format (ISO/IEC 26300, ODF) presentations from
Matlab.

Author: T Giorgino

Features
--------

- Cross-platform
- Writes arbitrary figures/diagrams (like _saveas_)
- Append to existing presentation, or replace
- Does not require MS Office to be installed
- Does not use COM or other IPC 
- It is based on Apache's POI (see installation instructions below), so assumes Matlab's JRE to be enabled 



Synopsis
--------

```
saveodp(h,filename,['-a'])
```

Saves the figure in the handle _h_ into the presentation _filename_. The figure is appended as a new slide if '-a' is passed as an argument. Otherwise, a single-slide presentation is created.


Example
-------

```
plot(1:10);
saveodp(gcf,'test.pptx');
plot(rand(10));
saveodp(gcf,'test.pptx','-a');
```



Installation
------------

The code uses Apache's odftoolkit library. You need to download and extract it in the same location as the script. Adjust the directory in the script if you have a different version than __0.6-incubator__. 

This matlab command may or may not suffice:

```
unzip('http://www.apache.org/dyn/closer.cgi/incubator/odftoolkit/binaries/odftoolkit-0.6-incubating-bin.zip')
```


Please note Apache odftoolkit's redistribution and licensing terms.
 


License
-------

    matlab_saveodp - Create Open Document Format presentations from Matlab. 
    Copyright (C) 2014 

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
 