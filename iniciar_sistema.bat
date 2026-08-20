@echo off
title Sistema de Fluxo de Caixa - Supermercado
echo Iniciando o sistema...
echo.
echo Depois de iniciar, acesse no navegador:
echo   Neste computador:  http://localhost:8000
echo   Em outros computadores da loja, use o IP deste PC, por exemplo: http://192.168.0.10:8000
echo   (rode "ipconfig" no cmd deste PC para descobrir o IP correto)
echo.
cd /d "%~dp0sistema"
py app.py
pause
