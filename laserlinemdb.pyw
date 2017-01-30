import win32com.client, pypyodbc, shutil, os, datetime, psutil, subprocess, numpy

PROCNAME = "blisk.exe"

for proc in psutil.process_iter():
    # check whether the process name matches
    if proc.name() == PROCNAME:
        proc.kill()
shutil.copy2('C:\\Users\\v.korolev\\Desktop\\dev\\python\\ntkdata.mdb', 'C:\\Users\\v.korolev\\Desktop\\dev\\python\\old')
target = "C:\\Users\\v.korolev\\Desktop\\dev\\python\\old\\"
file = "C:\\Users\\v.korolev\\Desktop\\dev\\python\\old\\ntkdata.mdb"
os.chdir(target)
allfiles = os.listdir(target)
for file in allfiles:
    v = datetime.datetime.now()
    x = v.strftime('%Y.%m.%d.')
os.rename('C:\\Users\\v.korolev\\Desktop\\dev\\python\\old\\ntkdata.mdb', (x +"ntkdata.mdb"))

conn = win32com.client.Dispatch(r'ADODB.Connection')
c = (
    r'PROVIDER=Microsoft.Jet.OLEDB.4.0;'
    r'DATA SOURCE=C:\\Users\\v.korolev\\Desktop\\dev\\python\\ntkdata.mdb;'
    )
conn.Open(c)
res = win32com.client.Dispatch(r'ADODB.Recordset')
res.Open("SELECT count(*) from history", conn)
data = res.GetRows()
int(data[0][0])



sres = win32com.client.Dispatch(r'ADODB.Recordset')
sres.Open("Select count(*) from history where logdate < now()-30", conn)
data2 = sres.GetRows()
int(data2[0][0])
z = data[0][0] - data2[0][0]

cmd = win32com.client.Dispatch(r'ADODB.Command')
if z > 100:
    cmd = win32com.client.Dispatch(r'ADODB.Command')
    cmd.ActiveConnection = conn
    cmd.CommandText = "DELETE FROM History where LogDate < now() - 30"
    cmd.Execute()
    
else:
    pass
conn.Close()
pypyodbc.win_compact_mdb("c:/users/v.korolev/desktop/dev/python/ntkdata.mdb","c:/users/v.korolev/desktop/dev/python/ntkdata.mdb")

subprocess.call("C:\\Users\\v.korolev\\AppData\\Local\\Blisk\\Application\\blisk.exe", shell=True)
