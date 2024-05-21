net stop spooler
del c:\windows\system32\spool\printers\*.shd
del c:\windows\system32\spool\printers\*.spl
net start spooler