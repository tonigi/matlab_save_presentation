Bmatlab_saveodp
==============

A command to create OpenOffice/LibreOffice presentations from Matlab. Please stand by while I populate the repository.

Features
--------

- Writes arbitrary figures/diagrams (like _saveas_)
- Append to existing presentation, or replace
- Does not require OpenOffice/LibreOffice to be installed
- Does not use COM or other IPC processes
- Output conforms to the OASIS document standard

Note that ODP presentations can be converted to PPT with the [unoconv](http://dag.wiee.rs/home-made/unoconv/) software (but it requires OO/LO to be installed).


Usage
-----

```
saveodp(h,filename,['-a'])
```

Saves the figure in the file handle _h_ into the OO presentation _filename_. The file is appended to if '-a' is passed as an argument.


Installation
------------

The software is composed by two parts: 
- the ```saveodp.m``` file, which you should copy anywhere in Matlab's path, and 
- the ```img2odp.pl``` script, which should be in the executable path. The script relies on a number of Perl modules to work. Installation is straightforward under Linux, less so under Windows.  For this reason,  a standalone-executable version is provided.

*Linux.* Use your distribution's package manager or the *cpan* command to install the _OpenOffice::OODoc_ module.
 

Troubleshooting
---------------

If you get...

```
Error using saveodp (line 13)
Can't locate OpenOffice/OODoc.pm in @INC (@INC contains: /home/toni/perl5/lib/perl5/i386-linux-thread-multi
/home/toni/perl5/lib/perl5/i386-linux-thread-multi /home/toni/perl5/lib/perl5 /usr/local/lib/perl5 /usr/local/share/perl5
/usr/lib/perl5/vendor_perl /usr/share/perl5/vendor_perl /usr/lib/perl5 /usr/share/perl5 .) at img2odp.pl line 13.
BEGIN failed--compilation aborted at img2odp.pl line 13.
```

...you need to install the OpenOffice::OODoc Perl module. Try e.g. to run ```cpan``` then ```install OpenOffice::OODoc```. Answer _yes_ to all questions.


License
-------

    matlab_saveodp - Create OpenOffice/LibreOffice presentations from Matlab.
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
