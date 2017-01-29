set shell = createobject("wscript.shell")
strtext = inputbox("What would you like the message to be? (TrapKilloYT)")
strtimes = inputbox("How many times would you like to type it? (TrapKilloYT)")
if not isnumeric(srtimes) then
lol=msgbox("Please write a NUMBER next time. (TrapKilloYT)")
wscript.quit
end if
strtdelay = inputbox ("What do you want the delay to be? (Milliseconds) (TrapKilloYT)")
msgbox "After you click ok, the message will start in 5 seconds. (TrapKilloYT)"
wscript.sleep(5000)
for i=1 to strtimes
shell.sendkeys(strtext & "")
shell.SendKeys "{Enter}"
wscript.sleep(strtdelay)
next
