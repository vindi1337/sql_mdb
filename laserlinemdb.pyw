import win32com.client, pypyodbc, shutil, os, datetime, psutil, subprocess

conn1 = win32com.client.Dispatch(r'ADODB.Connection')
c = (
    r'PROVIDER=Microsoft.Jet.OLEDB.4.0;'
    r'DATA SOURCE=C:\\LMC\\ConfigFiles\\ntkdata.mdb;'
    )
conn1.Open(c)
kres = win32com.client.Dispatch(r'ADODB.Recordset')
kres.Open("SELECT count(*) from history", conn1)
predata = kres.GetRows()
int(predata[0][0])

fres = win32com.client.Dispatch(r'ADODB.Recordset')
fres.Open("Select count(*) from history where logdate < now()-30", conn1)
afdata = fres.GetRows()
int(afdata[0][0])
x = predata[0][0] - afdata[0][0]

if x > 100:
    PROCNAME = "LaserMachine.exe"
    for proc in psutil.process_iter():
        if proc.name() == PROCNAME:
            proc.kill()
    file = "C:\\LMC\\ConfigFiles\\ntkdata.mdb"
    v = datetime.datetime.now()
    x = v.strftime('%Y.%d.%m.')
    shutil.copyfile(file, "C:\\LMC\\ConfigFiles\\old\\%sntkdata.mdb" % x)
    cmd = win32com.client.Dispatch(r'ADODB.Command')
    cmd.ActiveConnection = conn1
    cmd.CommandText = "DELETE FROM History where LogDate < now() - 30"
    cmd.Execute()
    conn1.Close()
    pypyodbc.win_compact_mdb("c:/LMC/ConfigFiles/ntkdata.mdb","c:/LMC/ConfigFiles/ntkdata.mdb")
    subprocess.call("C:\\Program Files\\Nutek\\LMC\\LaserMachine.exe", shell=True)
else:
    pass
