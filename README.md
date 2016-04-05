# Archives Finder
archives_finder.vbs

DESCRIPTION:
The objective of this script is to allow archivists to find groups of records
that may be inactive because of their age.  It is designed to be run across
large networked file systems, although it can be run across any storage device.  
The software finds the largest possible groupings of folders with files in them
that are a given number of years old, based on the date last modified attribute.  
Since this atribute can be easily and accidentally modified (e.g., someone opening
a file and saving it), the program allows for some fuzzy math: 
allowing a threshold where some percentage of files must be X years old.  
The default threshold is 95%, but this can be adjusted as needed.  The default
number of years is seven, and this can be adjusted as well, and decimal years
(e.g., 7.5) can be used to indicate portions of a year.

Written in VB Script.  Tested on Windows XP and Windows 7.

LICENSE:
Copyright Anthony Cocciolo
This software is licensed under the CC BY-NC-SA 3.0 license
see: http://creativecommons.org/licenses/by-nc-sa/3.0/
