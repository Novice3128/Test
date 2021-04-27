@ECHO off


for /F "tokens=*" %%1 in ('wevtutil el') DO wevtutil cl "%%1"
EXIT

