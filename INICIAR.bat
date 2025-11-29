@echo off
:: Este comando abre o PowerShell, ignora a política de segurança e roda o script .ps1 da mesma pasta
PowerShell.exe -ExecutionPolicy Bypass -File "%~dp0Conversor.ps1"