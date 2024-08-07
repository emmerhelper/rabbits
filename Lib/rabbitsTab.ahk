#Include utils.ahk
Class rabbitsTab extends Object {
      
      __New(rabbits,parentTab,components) {

            rabbits.tabs.UseTab(parentTab)
            this.parentTab := parentTab
            this.Type := rabbits.tabnames[parentTab]
            this.plural := getPlural()
            this.components := components
           
            if components.Generation
                  AddGenerationSection()

            if components.People
                  AddPeopleSection()

            if components.Processing
                  AddProcessingSection()

            
            getPlural(){
                  if (this.type = "Staff"){
                        return this.type
                  } else return this.type "s"
            }

            addGenerationSection(){
                  rabbits.mainWindow.SetFont("w600 s12")
                  rabbits.mainWindow.AddText("Section","Generation")
                  rabbits.mainWindow.SetFont("w100")
                  rabbits.mainWindow.AddText(,"Nationality: ")
                  
                  this.nationalityDropDown := rabbits.mainWindow.AddDropDownList("w182 choose" . rabbits.ini.Nationality . " vNationality" this.parentTab,getNationalities())
                  
                  rabbits.mainWindow.AddText(,"Amount to generate: ")
                  rabbits.mainWindow.Add("Edit", "w182")
                  
                  this.count := rabbits.mainWindow.Add("UpDown", "w182 range1-160 vCount" this.parentTab, rabbits.ini.Amount_To_Generate)
                  
                  rabbits.mainWindow.AddText(,"Process new " . StrLower(this.plural) ": ")
                  
                  this.processAfterCheckbox := rabbits.mainWindow.AddCheckbox("x+1 vafterGeneration" this.parentTab)
                  
                  if !this.components.Processing{
                        this.processAfterCheckbox.Enabled := false
                  }

                  this.generateButton := rabbits.mainWindow.AddButton("xs y+m center w182 h65","Generate")
           
                  this.generateButton.onEvent("Click",runProcesses)
                  this.processAfterCheckbox.OnEvent("Click",disableProcessButton)
            }
            
            AddPeopleSection(){
                  rabbits.mainWindow.SetFont("w600")
                  rabbits.mainWindow.AddText("ys x+m",this.plural)
                  rabbits.mainWindow.SetFont("w100")
                  
                  this.numbers := rabbits.mainWindow.AddEdit("r11 w182 -wrap")
            }
            
            addProcessingSection(){
                  rabbits.mainWindow.SetFont("w600")
                  rabbits.mainWindow.AddText("ys x+m Section","Processing")
                  rabbits.mainWindow.SetFont("w100")  
                  this.action := rabbits.mainWindow.AddDropDownList("choose1 vAction" this.parentTab,getActions())
                  this.addMyButton := rabbits.mainWindow.AddButton("x+10 w120 hp","Add")
                  this.listView := rabbits.mainWindow.AddListView("xs r6 checked w310",["Command","Order","Value","Max"])
                  this.listView.ModifyCol(1, "AutoHdr")
                  this.processButton := rabbits.mainWindow.AddButton("center w310 h28","Process")
                  this.processButton.onEvent("Click",runProcesses)
                  this.listView.OnEvent("Click", listViewChangeValue)
                  this.listView.OnEvent("ContextMenu", listViewDelete)
                  this.listView.OnEvent("ItemCheck", listViewSortChange)
                  this.addMyButton.OnEvent("Click", addToListView)
            }

            getActions(){
                  switch this.Type {
                        case "Student":
                              actions := ["Register","Admit","ZPIQSU01","Set_Home_Student","Link_To_Advisor","Link_To_Cohort"]
                        case "Staff":
                              actions := ["Link_to_module","Delete"]
                        case "Module":
                              actions := ["Link_to_staff"]
                        Default:
                              actions := []
                  }
                  return actions
            }
            
            addToListView(params*){
                  try {
                        parameter1 := iniRead("config.ini",this.action.text,"value")
                  } catch {
                        parameter1 := ""
                  }
                  
                  if !this.action.text{
                        return 
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
                  
                  this.listView.Add("check",this.action.text,sortOrder,parameter1) 
                  this.listView.ModifyCol(1, "Auto")
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
            
            listViewChangeValue(obj, rowNumber){
                  if !rowNumber
                        return 
                  
                  MouseGetPos(&mouseX,&mouseY)

                  rabbits.mainWindow.Opt("Disabled")

                  listItemEditMenu := Gui(,this.ListView.GetText(rowNumber,1))
                  listItemEditMenu.SetFont("s12 w600")
                  listItemEditMenu.AddText(,"Order")
                  listItemEditMenu.SetFont("w100")
                  listItemEditMenu.Add("Edit")
                  order := listItemEditMenu.Add("UpDown","w180 range1-100",this.ListView.GetText(rowNumber,2))
                  listItemEditMenu.SetFont("s12 w600")
                  listItemEditMenu.AddText("xs","Value")
                  listItemEditMenu.SetFont("w100")
                  value := listItemEditMenu.Add("Edit","w180",this.ListView.GetText(rowNumber,3))
                  listItemEditMenu.SetFont("s12 w600")
                  listItemEditMenu.AddText("xs","Continue after")
                  listItemEditMenu.SetFont("w100")
                  HP := listItemEditMenu.Add("Edit","w180 Number",this.ListView.GetText(rowNumber,4))
                  saveButton := listItemEditMenu.AddButton("xs w180","Save")
                  saveButton.onEvent("Click",saveListItemEdits)
                  listItemEditMenu.OnEvent('Close',closeListItemEdits)
                  listItemEditMenu.OnEvent('Escape',closeListItemEdits)


                  listItemEditMenu.Show()
                  Send("{Tab}")

                  saveListItemEdits(obj,info){
                        this.listView.Modify(rowNumber,"col2",order.Value)
                        this.listView.Modify(rowNumber,"col3",value.Value)
                        this.listView.Modify(rowNumber,"col4",HP.Value)


                        closeListItemEdits()
                  }

                  closeListItemEdits(*){
                        listItemEditMenu.Destroy()
                        rabbits.mainWindow.Opt("-Disabled")
                        WinActivate(rabbits.mainWindow)
                  }
            
            }
            
            runProcesses(button,info){
                  ;; Run all ticked commands in the list view, in order, generating first if needed
                  getSessionInfo()
                  
                  if (button.text = "Generate"){
                        resolveCommand("create",this.nationalityDropDown.Text)
                        if (!this.processAfterCheckbox.value){
                              return 
                        }      
                  }

                  if !this.numbers.Value{
                        Msgbox "Add some numbers first."
                        return 
                  }
                  
                  RowNumber := 0

                  loop {
                        RowNumber := this.listView.GetNext(RowNumber,"C")
                        
                        if !RowNumber{
                              break
                        }
            
                        resolveCommand(this.listView.GetText(RowNumber),this.listView.GetText(RowNumber,3),this.listView.GetText(RowNumber,4))                  
                  }
            }
      
            getSessionInfo(){
                  try 
                        openMaxSessions(rabbits.sessions.value)
                  catch {
                        noSAP()
                        Exit
                  }
                  rabbits.SystemName := sapActiveSession().info.SystemName
            }
               
            disableProcessButton(*){
                  ;; For when we directly run commands on new students
                  
                  if (this.processAfterCheckbox.value = 1)
                        this.processButton.Opt("+Disabled")
                  else this.processButton.Opt("-Disabled")
                  
            }
            
            resolveCommand(command,param7,HP:=0){
                  ;; Run the command on each student in as many parallel sessions as we're allowed
                  if HP {
                        hpFlag := true
                  } else hpFlag := false

                  originalHP := HP

                  ;; Build an array we can iterate over
                  if (command = "Create"){
                        ;; Clear the edit box
                        this.Numbers.value := ""
                        
                        people := []
                        loop this.count.value {
                              people.Push(param7)
                        }

                        rabbits.AddReportSection()
                  } else {
                        people := StrSplit(Trim(this.numbers.value,"`r`n"),"`n") 
                  }
                  
                  session := 0
                  
                  for k, Number in people {
                     ;; Arguments:
                        ;; A_Args[1] is the system name that sap is connecting to
                        ;; A_Args[2] is the session of SAP that this instance controls
                        ;; A_Args[3] is the count (how many have already been generated)
                        ;; A_Args[4] is the type to generate
                        ;; A_Args[5] is the process to be run
                        ;; A_Args[6] is number or nationality
                        ;; A_Args[7] is a command parameter, - the value of the command
                        
                        Run "worker.ahk " rabbits.SystemName " " session " " A_Index " " this.Type " " command " " number " " param7
                        session++
                        WinWait("Rabbits are working, please wait warmly 🐇")
                        
                        if (session = rabbits.sessions.value){
                              WinWaitClose("Rabbits are working, please wait warmly 🐇")
                              session := 0
                        }

                        if hpFlag {
                              HP--
                              if HP = 0
                                    break
                        }
                  }     

                  if !hpFlag 
                        return

                  people.RemoveAt(1,originalHP)
                  newPeopleString := ""
                  for k, people in people {
                        newPeopleString .= people "`r`n"
                  }
                  this.numbers.value := newPeopleString
            } 
      }
}

#HotIf WinActive("")

