Dim Message, Speak
Message=InputBox("Masukan Kalimat","Riventus.tech")
Set Speak=CreateObject("sapi.spvoice")
Speak.Speak Message