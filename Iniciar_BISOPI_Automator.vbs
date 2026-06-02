Set oShell = CreateObject("WScript.Shell")
oShell.CurrentDirectory = "C:\BISOPI_Automator"
oShell.Run """C:\Users\Usuario\AppData\Local\Programs\Python\Python310\python.exe"" -m streamlit run main.py", 0, False
WScript.Sleep 3000
oShell.Run "http://localhost:8501", 1, False
