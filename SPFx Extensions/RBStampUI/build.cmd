@echo off
cls
call npx gulp clean
call npx gulp bundle --ship
call npx gulp package-solution --ship
call explorer .\sharepoint\solution\

