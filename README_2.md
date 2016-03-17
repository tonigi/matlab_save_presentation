matlab_savepptx
===============

Create Office Open XML presentations from Matlab.  

Author: T Giorgino
        ISIB-CNR

Features
--------

- Cross-platform
- Writes arbitrary figures/diagrams (like _saveas_)
- Append to existing presentation, or replace
- Does not require MS Office to be installed
- Does not use COM or other IPC 
- It is based on Apache's POI (see installation instructions below), so assumes Matlab's JRE to be enabled 

Note that PPTX presentations can be opened in LibreOffice and in Office 2003 versions with the MS Compatibility pack.

This code was written before discovering the [exportToPPTX](https://github.com/stefslon/exportToPPTX) pure-Matlab implementation, which offers much more flexibility. You may find, however, _savepptx_'s syntax more straightforward and that it produces higher-resolution images by default. 



Synopsis
--------

```
savepptx(h,filename,['-a'])
```

Saves the figure in the handle _h_ into the PPTX presentation _filename_. The figure is appended as a new slide if '-a' is passed as an argument. Otherwise, a single-slide presentation is created.


Example
-------

```
plot(1:10);
savepptx(gcf,'test.pptx');
plot(rand(10));
savepptx(gcf,'test.pptx','-a');
```



Installation
------------

The code uses Apache's POI library. You need to download Apache's [POI binary distribution](http://poi.apache.org/download.html) and extract it in the same location as the script. Adjust the directory in the script if you have a different version than __3.10-FINAL__. 

This matlab command may or may not suffice:

```
untar('http://apache.fastbull.org/poi/release/bin/poi-bin-3.10-FINAL-20140208.tar.gz')
```


Please note Apache POI's redistribution and licensing terms.
 


License
-------

    matlab_savepptx - Create Office Open XML presentations from Matlab. 
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
 