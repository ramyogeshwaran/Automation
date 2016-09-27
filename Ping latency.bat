@ECHO OFF

TITLE Ping latency
echo checking the network latency

Rem Settings
SET "IP=RZ2DCVPZ02V"            %= server name or IP  address %
SET "LOG=%IP%.log"              %= Log File name and Location %
SET "LTime=150"                 %= If OutResult Exceeds this time in ms it will turn the screen RED =%

:Loop

Rem Ping The IP One Time

set "var="
For /F "delims=" %%A IN (' Ping "%IP%" -n 1 ^|find "Reply from" ') DO (
      >>"%LOG%" Echo %date% %time% - %%A
      set "var=%%A"
)

Rem if the site does not respond then indicate that

if not defined var (
echo %date% %time% - "%IP%" not responding
>>"%log%" echo %date% %time% - "%IP%" not responding
goto :loop
)

Rem Get the time result

For /F "tokens=5 delims==m" %%A In ("%var%") Do SET "OutResult=%%A"

   Rem Allowed Range "green"

   IF %OutResult% LEQ %LTime% color a0

   Rem Exceeded Allowed range "red"

   IF %OutResult% GTR %LTime% color c0

   echo %date% %time%: IP %var:~11%

start cscript c:\mail.vbs

Rem delay for 10 seconds

ping -n 10 localhost >nul

Goto :LOOP