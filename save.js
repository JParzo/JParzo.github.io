function WriteToFile(passForm) {
 
    set fso = CreateObject("Scripting.FileSystemObject"); 
    set s   = fso.CreateTextFile("<your Path>/filename.txt", True);
 
    var name = document.getElementById('Name');
 
    s.writeline(name);
    s.writeline("-----------------------------");
    
    s.Close();
 }
