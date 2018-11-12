on error resume next
url = "https://webmail30.189.cn/w2/"
username = "jixing_xie"
password = ""
set ie = CreateObject("InternetExplorer.Application")
ie.visible = true
ie.Navigate url,4 ' " http://www.baidu.com"
do until 4=ie.readyState
    WScript.sleep 200
    waittime  = waittime + 200
    if waittime > 15000 then exit do
loop
'WScript.echo waittime
if 4<>ie.readyState then
    ie.quit
    WScript.quit
end if
set dom = ie.document

set user = dom.getElementById("userName")
set pswd = dom.getElementById("password")
user.value = username
pswd.value = password
WScript.echo user.value

' set form = dom.getElementById("iframeLogin")
' form.all("userName").value = username
' form.all("password").value = password
' form.all("j-login").click()
