DEFINE g_etapp integer
DEFINE g_etwb  integer

MAIN
  DEFINE l_result integer
  DEFINE str string
  DEFINE l_count integer
  DEFINE l_etwb string

  LET g_etapp=-1
  LET g_etwb=-1

  CALL ui.Interface.frontCall("WinCOM","CreateInstance",["ET.Application"],[g_etapp])
  CALL CheckError(g_etapp, __LINE__)

  CALL ui.Interface.frontCall("WinCOM","CallMethod",[g_etapp,"WorkBooks.add"],[g_etwb])
  CALL CheckError(g_etwb,__LINE__)

  CALL ui.Interface.frontCall("WinCOM","SetProperty",[g_etapp,"Visible",true],[l_result])
  CALL CheckError(l_result,__LINE__)

  CALL ui.Interface.frontCall("WinCOM","SetProperty",[g_etwb,'activesheet.Range("A1:B3").Value',"hello"],[l_result])
  CALL CheckError(l_result,__LINE__)

   CALL ui.Interface.frontCall("WinCOM","GetProperty",[g_etwb,'activesheet.Range("A1:B3").Value'],[str])
   CALL CheckError(str,__LINE__)

   CALL ui.Interface.frontCall("WinCOM","GetProperty",[g_etwb,'Worksheets.Count'],[l_count])
   CALL CheckError(l_count,__LINE__)

   CALL ui.Interface.frontCall("WinCOM","SetProperty",[g_etwb,'Worksheets.Item(2).Range("A1:B3").Value',"world"],[l_result])
   CALL CheckError(l_result,__LINE__)

   LET l_etwb="E:\test.xls"
  CALL ui.Interface.frontCall("WinCOM","CallMethod",[g_etapp,"WorkBooks.open"],[l_etwb])
  CALL CheckError(l_etwb,__LINE__)
    
   DISPLAY "content of the cell is: " || str ,l_count
END MAIN

FUNCTION FreeMemory()
   DEFINE l_res INTEGER
   IF g_etwb != -1 THEN
     CALL ui.Interface.frontCall("WinCOM","ReleaseInstance", [g_etwb], [l_res] )
   END IF
   IF g_etapp != -1 THEN
     CALL ui.Interface.frontCall("WinCOM","ReleaseInstance", [g_etapp], [l_res] )
   END IF
END FUNCTION

FUNCTION CheckError(p_res,p_lin)
DEFINE p_res integer,
       p_lin integer,
       l_mes string
    IF p_res=-1 THEN
       DISPLAY "COM Error for call at line:",p_lin
       CALL ui.Interface.frontCall("WinCom","GetError",[],[l_mes])
       DISPLAY "Error:",l_mes
       CALL FreeMemory()
       DISPLAY "Exit with COM Error."
       EXIT PROGRAM (-1)
   END IF
END FUNCTION

