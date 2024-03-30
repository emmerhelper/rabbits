

lineFromFile(country,path){
      ;; Returns a random line from the specified path
      
            lines := StrSplit(FileRead("./Nationalities/" . country . "/" path . ".txt"),"`n")
      
            return Trim(lines[Random(1, lines.Length)],"`r")
      
      }

generateLetters(length){
      ;; Return a string of random uppercase characters of the requested length 
      
            string := ""
      
            loop length {
      
                  string .= Chr(Random(65,90))
      
            }
      
            return string
      
      }


sapConnect(systemName:=false,instance:=0){
      ;; Return a specific session of a specific system. If no system is specified, it returns the first child.

            SAP := ComObjGet("SAPGUI")
      
            app := SAP.GetScriptingEngine()

            
            if !systemName {
                  connection := app.children[0]
                  session := false 
                  while !session {
                        session := connection.children[Integer(instance)]
                  }
                  return session
            }

            for k, v in app.children{

                  connection := app.children[Integer(A_Index-1)]

                  if (connection.children[0].info.systemName = systemName)
                        return connection.children[Integer(instance)]
                  
            } 

          
      }


      sapActiveSession(){
            ;; Activates the last used SAP window and returns it as a session.
                  
                  WinActivate("ahk_exe saplogon.exe")
            
                  SAP := ComObjGet("SAPGUI")
                  app := SAP.GetScriptingEngine()
            
                  return app.ActiveSession()
            }
      
            

      openMaxSessions(max){
            ;; Opens the max number of sessions allowed, and tells us what kind of system it is 
                  WinActivate("ahk_exe saplogon.exe")
                  SAP := ComObjGet("SAPGUI")
                  app := SAP.GetScriptingEngine()
                  
                  for k, v in app.children{

                        connection := app.children[Integer(A_Index-1)]
      
                        if (connection.children[0].info.systemName = app.activeSession().info.systemName){
                              session := connection.children[Integer(0)]
                              break 
                        }
                  } 
            

                  while (connection.children.length < max){
                              session.createSession()
                              sleep 500
                        }

            }
            
      getNationalities(){
            ;; Read all two character folders in the working directory and return them as an array.
            
                  nationalities := []
            
                  loop files, A_WorkingDir "./Nationalities/*", "D R" {
            
                        nationalities.Push(A_LoopFileName)
            
                  }
            
                        
            
                  return nationalities
            }


                  
      noSAP(){
            ;; Handle not having SAP open.
            
                  Msgbox("No instance of SAP was found. Check if GUI scripting is enabled on the server and in your user settings.","Rabbits")
            }

      findTextElement(userArea,name){
            ;;Checks for both Text and Ctext
            element := false
            
            while !element {
                  if element := userArea.findByName(name,"GuiTextField")
                        return element
                  if element := userArea.findByName(name,"GuiCTextField")
                        return element 
            }
      
      }

      selectStudentFileTab(userArea,name){
            
            ;; Tab = Student File, DETLATAB = Student Master Data 
            if userArea.findByName("TAB01", "GuiTab"){
                  tabPrefix := "TAB"
            } else if userArea.findByName("DETLTAB01","GuiTab"){
                  tabPrefix := "DETLTAB"
            }

            ;; Look for the tab with the requested name
            while true {
                  tabNumber := Format("{:02}",A_Index)
                  tab := userArea.findByName(tabPrefix . tabNumber,"GuiTab")             
                  if tab.Text = name {
                        tab.select()
                        return 
                  }
            }
      }

      openStudentInStudentFile(session, studentNumber){

            session.startTransaction("PIQST00")
            userArea := session.findByID("wnd[0]/usr")
            findTextElement(userArea,"PIQST00-STUDENT12").Text := studentNumber
            session.findById("wnd[0]").sendVKey(0)
            return userArea

      }

showAllParams(array){
      ;; Nice gui for debug
      debug := Gui(,"Props")
      for k, v in array.OwnProps(){
            debug.AddText(,k ": " v)
      }
      debug.Show
      msgbox "showing debug"

}

returnIniSectionAsObject(ini,sectionName){

      section := {}
      sectionFound := false

      loop read ini ".ini" {

            if !A_LoopReadLine 
                  continue 

            if InStr(A_LoopReadLine,";")
                  continue

            if !sectionFound {
                  if (A_LoopReadLine = "[" sectionName "]")
                        sectionFound := true
                  continue 
            }

            if InStr(A_LoopReadLine,"[")
                  break

            values := StrSplit(A_LoopReadLine,"=")
            section.%values[1]% := values[2]
      }

      return section


}