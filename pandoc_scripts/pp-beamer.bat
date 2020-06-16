@echo off


::filename
set par1=%1
::output
set par2=%2

IF ["%par1:~-2%"] == ["md"] ( set fname=%par1:~0,-3%)

IF ["%par2%"] == ["pdf"] ( 
pandoc -s -t beamer -V theme:Madrid -V colortheme:beaver --slide-level=2 %par1% -o %fname%.%par2%)

IF ["%par2%"] == ["tex"] ( 
pandoc -s -t beamer -V theme:Madrid -V colortheme:beaver --slide-level=2 %par1% -o %fname%.%par2%)

IF ["%par2%"] == ["pptx"] ( 
pandoc -s --slide-level=2 %par1% -o %fname%.%par2%)