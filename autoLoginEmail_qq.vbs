on error resume next
url = "https://mail.qq.com/"
username = "1327183688"
password = ""
set ie = CreateObject("InternetExplorer.Application")
ie.visible = true
ie.Navigate url,4
do until 4=ie.readyState
    WScript.sleep 200
    waittime  = waittime + 200
    if waittime > 15000 then exit do
loop
'WScript.echo waittime
if 4<>ie.readyState then
	WScript.echo "i am here"
    ie.quit
    WScript.quit
end if
set dom = ie.document
'set form = dom.getElementById("loginform")
set user = dom.getElementById("u")
set pswd = dom.getElementById("p")
WScript.echo user.value
'user.value = username
'pswd.value = password
'form.all("u").value = username
'form.all("p").value = password
'form.all("login_button").click()
