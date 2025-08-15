@ftp -i -s:"%~f0"&GOTO:EOF
open 10.120.1.42
kwiintf
intfKwi1
binary
lcd "C:\EZRunner\IFData"
cd "/ki"
mput *.csv
disconnect
quit
