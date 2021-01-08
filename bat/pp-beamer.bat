@echo off


::filename
set par1=%1
::output
set par2=%2

IF ["%par1:~-2%"] == ["md"] ( set fname=%par1:~0,-3%)

IF ["%par2%"] == ["pdf"] ( 
pandoc %~dp0\header-includes.yaml -s -t beamer --pdf-engine=xelatex --slide-level=2 %par1% -o %fname%.%par2%)

IF ["%par2%"] == ["tex"] ( 
pandoc %~dp0\header-includes.yaml -s -t beamer --pdf-engine=xelatex --slide-level=2 %par1% -o %fname%.%par2%)

IF ["%par2%"] == ["pptx"] ( 
pandoc -s --slide-level=2 %par1% -o %fname%.%par2%)