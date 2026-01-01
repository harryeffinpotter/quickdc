@echo off
powershell -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0reset-wifi.ps1\"' -Verb RunAs"
