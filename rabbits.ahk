;Rabbits v1.0
#Requires AutoHotkey v2.0 
#Include lib\utils.ahk
#SingleInstance Force

TraySetIcon("icon_jc2_icon.ico")
Mine := MainGUI()

Class MainGUI extends Gui {


      __New(){
            ;;Build GUI
            mainWindow := Gui(,"Rabbits")
            readIniDefaults()

            AddGenerationSection()
            AddStudentsSection()
            AddProcessingSection()
            AddOptionsSection()
            SetEventListeners()
              
            mainWindow.Show()    


      readIniDefaults(){
            this.ini := {}
            this.ini.Nationality := IniRead("config.ini","GUI_Defaults","Nationality")
            this.ini.Amount_To_Generate := IniRead("config.ini","GUI_Defaults","Amount_To_Generate")
            this.ini.Max_Sessions := IniRead("config.ini","GUI_Defaults","Max_Sessions")
      }

      addGenerationSection(){
            mainWindow.SetFont("w600 s12")
            mainWindow.AddText(,"Generation")
            mainWindow.SetFont("w100")
            mainWindow.AddText(,"Nationality: ")
            this.nationality := mainWindow.AddDropDownList("w182 choose" . this.ini.Nationality . " vNationality",getNationalities())
            mainWindow.AddText(,"Amount to generate: ")
            mainWindow.Add("Edit", "w182")
            
            count := mainWindow.Add("UpDown", "w182 vCount range1-160", this.ini.Amount_To_Generate)
            mainWindow.AddText(,"Process new students: ")

            this.processStudentsAfter := mainWindow.AddCheckbox("x+1 vafterGeneration")

            this.generateButton := mainWindow.AddButton("xs center w182 h60","Generate")
      }

      addStudentsSection(){
            mainWindow.SetFont("w600")
            mainWindow.AddText("ym","Students")
            mainWindow.SetFont("w100")
            
            this.studentNumbers := mainWindow.AddEdit("r11 w182 number -wrap")
            
      }

      addProcessingSection(){
            mainWindow.SetFont("w600")
            mainWindow.AddText("ym section","Processing")
            mainWindow.SetFont("w100")
            
            this.action := mainWindow.AddDropDownList("xs choose1 vAction",getActions())
            this.addMyButton := mainWindow.AddButton("x+10 w120","Add")


            this.listView := mainWindow.AddListView("xs r5 checked w310",["Process","Order","Value"])

            this.listView.ModifyCol(1, "AutoHdr")


            this.processButton := mainWindow.AddButton("center w310","Process")
            
      }

      addOptionsSection(){

            mainWindow.SetFont("w600")
            mainWindow.AddText("ym section","Options")
            mainWindow.SetFont("w100")


            mainWindow.AddText(,"Max sessions: ")
            mainWindow.Add("Edit", "w182")
            sessions := mainWindow.Add("UpDown", "w182 vsessions range1-16", this.ini.Max_Sessions)
            

            
            mainWindow.AddText("xs","Open configuration file: ")
            this.SettingsButton := mainWindow.AddButton("w182","Settings")
 


      }

      AddReportSection(){

            if WinExist("Rabbits Report"){
                  return 
            }
            
            this.reportWindow := Gui(,"Rabbits Report")
            this.reportWindow.SetFont("w600")
            this.reportHeader := this.reportWindow.AddText("ym","Report")
            this.reportWindow.SetFont("w100")
            this.Report := this.reportWindow.AddEdit("r9 w300")
            this.ExportButton := this.reportWindow.addButton(,"Export")
            this.ExportButton.onEvent("Click",exportToCSV)

            this.reportWindow.Show()

      }

      setEventListeners(){
      ;; Event handling
      this.generateButton.onEvent("Click",newStudents)
      this.SettingsButton.onEvent("Click",settingsFile)
      this.processButton.onEvent("Click",runProcesses)
      this.processStudentsAfter.OnEvent("Click",disableProcessButton)
      this.listView.OnEvent("DoubleClick", listViewChangeValue)
      this.listView.OnEvent("ContextMenu", listViewDelete)

      this.listView.OnEvent("ItemCheck", listViewSortChange)
      this.addMyButton.OnEvent("Click", addToListView)

      mainWindow.OnEvent("Close",Quit)

      }


      
      quit(*){
      ;; Close the app when we close the window
            ExitApp
      }

      addToListView(params*){

            this.contents := mainWindow.Submit(false)
            try {
                  parameter1 := iniRead("config.ini",this.contents.action,"value")
            } catch {
                  parameter1 := ""
            }

            sortOrder := 1
            rowNumber := 0

            loop {

                  RowNumber := this.listView.GetNext(RowNumber,"C")
                  
                  if !RowNumber{
                        break
                  }

                  else sortOrder := this.ListView.GetText(RowNumber,2) + 1

            }

            this.listView.Add("check",this.contents.action,sortOrder,parameter1) 
            this.listView.ModifyCol(1, "Auto")



      }


      exportToCSV(params*){

            Export := "Student Number,First Name,Last Name,Nationality`r`n" . this.Report.Value 
            Path := FileSelect("S 16","Generated Students.csv")
            try {
                  FileRead(path)
                  FileDelete(path)
            }
            FileAppend(export,Path,)
            Run(path)
            this.reportWindow.hide()

      }

      listViewSortChange(params*){
      ;; Check the checkbox to change the sorting value of that row

            ;; Reset the sorting when the checkbox is unchecked
            if (params.Length = 3){
                  if params[3] = 0{
                        this.listView.Modify(params[2],"col2","❌")
                        return 
                  }
            }

            ;; Ask for a new sort on checking/double click
            input := InputBox(,"Enter sort order")
            
            if (input.result = "OK"){
                  this.listView.Modify(params[2],"col2",input.Value)
            }
            else this.listView.Modify(params[2],"col2","❌")
                  
            this.listView.ModifyCol(2, "AutoHdr")
            this.listView.ModifyCol(2, "Sort")
            
      }

      listViewDelete(obj, item, isRightClick, X, Y){
            try obj.delete(item)      
      }

      listViewChangeValue(obj, info){

            input := InputBox(,"Enter new value")
            
            if (input.result = "OK"){
                  this.listView.Modify(info,"col3",input.Value)
            }

      }


      settingsFile(*){

            RunWait("config.ini")

      }

            
      disableProcessButton(*){
      ;; For when we directly run commands on new students

            if (this.processStudentsAfter.value = 1)
                  this.processButton.Opt("+Disabled")
            else this.processButton.Opt("-Disabled")

      }
      
      runProcesses(*){
      ;; Run all ticked commands in the list view, in order

            this.contents := mainWindow.Submit(false)

            getSessionInfo()

            if !this.studentnumbers.Value{
                  Msgbox "Add some student numbers first."
                  return 
            }

            RowNumber := 0

            loop {

                  RowNumber := this.listView.GetNext(RowNumber,"C")
                  
                  if !RowNumber{
                        break
                  }

                  else resolveCommand(this.listView.GetText(RowNumber),this.listView.GetText(RowNumber,3))

            }

      }

      resolveCommand(command,param5){
      ;; Run the command on each student in as many parallel sessions as we're allowed

            studentNumberString := StrSplit(Trim(this.studentNumbers.value,"`r`n"),"`n") 
            
            counter := 0

            for k, studentNumber in studentNumberString {
                  
                  ;; Parameters:
                  ;; 1 = System name
                  ;; 2 = Session to connect to (Counter)
                  ;; 3 = Current student (How many students we've processed thus far, so A_Index)
                  ;; 4 = Command (Command to run)
                  ;; 5 = Function parameter (Student number)
                  ;; 6 = Function parameter (Value)

                  Run "worker.ahk" . " " this.SystemName . " " . counter . " " . A_Index . " " . command . " " . studentNumber . " " . param5
                  counter++
                  WinWait("Rabbits are working, please wait warmly 🐇")
                
                  if (counter = this.contents.sessions){
                        WinWaitClose("Rabbits are working, please wait warmly 🐇")
                        counter := 0
                  }
                       
                  
                  
            }           



      }

      getSessionInfo(){
            try 
                  openMaxSessions(this.contents.sessions)
            catch {
                  noSAP()
                  Exit
            }

            this.SystemName := sapActiveSession().info.SystemName
      
      }

      newStudents(*){
      ;; Generate as many students as requested

            this.contents := mainWindow.Submit(false)

            this.studentNumbers.value := ""

            getSessionInfo()
            
            AddReportSection()

     

            counter := 0

            loop this.contents.count {
                  
                  ;; Parameters:
                  ;; 1 = System name
                  ;; 2 = Session
                  ;; 3 = Count
                  ;; 4 = Command
                  ;; 5 = Function parameter (nationality)
                  Run "worker.ahk" . " " . this.SystemName . " " . counter . " " . A_Index . " Generate " . this.contents.nationality
                  
                  ++counter
                  WinWait("Rabbits are working, please wait warmly 🐇")

                 
                  if (counter = this.contents.sessions){
                        WinWaitClose("Rabbits are working, please wait warmly 🐇")
                        counter := 0
                  }
                       
            }
            
            

            if !this.processButton.enabled{
                  sleep 1000
                  runProcesses()
            }
      }
}
}
