;Rabbits v1.0
#Requires AutoHotkey v2.0 
#Include lib\utils.ahk
#Include lib\rabbitsTab.ahk
#SingleInstance Force

TraySetIcon("icon_jc2_icon.ico")
Mine := MainGUI()

Class MainGUI extends Gui {

      __New(){

            ;;Build GUI
            this.mainWindow := Gui(,"Rabbits")
            this.ini := returnIniSectionAsObject("config","GUI_Defaults")
            this.tabnames := ["Student","Staff","Module","Options"]
            this.tabs := this.mainWindow.AddTab3("Backgroundffffff",this.tabnames)
            this.tabs.onEvent("Change",resizeWindow)
            this.mainWindow.OnEvent("Close",Quit)

            this.studentTab := rabbitsTab(this,1,{Generation: true, People: true, Processing: true})
            this.staffTab := rabbitsTab(this,2,{Generation: true, People: true, Processing: false})
            this.moduleTab := rabbitsTab(this,3,{Generation: false, People: true, Processing: true})

            addOptionsSection()
            
            this.mainWindow.Show()    

            addOptionsSection(){
                  this.tabs.UseTab(this.tabnames.Length)
                  this.mainWindow.SetFont("w600")
                  this.mainWindow.AddText("ys x+m section","Options")
                  this.mainWindow.SetFont("w100")
                  this.mainWindow.AddText(,"Max sessions: ")
                  this.mainWindow.Add("Edit", "w182")
                  this.sessions := this.mainWindow.Add("UpDown", "w182 vsessions range1-16", this.ini.Max_Sessions)
                  this.mainWindow.AddText("xs","Open configuration file: ")
                  this.SettingsButton := this.mainWindow.AddButton("w182","Settings")
                  this.SettingsButton.onEvent("Click",settingsFile)
            }

            quit(*){
            ;; Close the app when we close the window
                  ExitApp
            }

            resizeWindow(Tab,Info){

                  switch this.tabnames[tab.value] {
                        case "Student":
                              this.mainWindow.Move(,,746,348)
                        case "Staff":
                              this.mainWindow.Move(,,425,348)
                        case "Module":
                              this.mainWindow.Move(,,560,348)
                        case "Options":
                              this.mainWindow.Move(,,250,250)
                  }
            }

            settingsFile(*){
                  RunWait("config.ini")
            }
      }
      
      AddReportSection(){

            if WinExist("Rabbits Report"){
                  return 
            }
            
            this.reportWindow := Gui(,"Rabbits Report")
            this.reportWindow.SetFont("w600")
            this.reportHeader := this.reportWindow.AddText("ym","Report")
            this.reportWindow.SetFont("w100")
            this.Report := this.reportWindow.AddEdit("r9 w1000")
            this.ExportButton := this.reportWindow.addButton(,"Export")
            this.ExportButton.onEvent("Click",exportToCSV)
            this.reportWindow.Show("Minimize")

            exportToCSV(params*){

                  Export := "City,Email,House Number,Phone Number,Postcode,Street,Nationality,Number,Birthplace,Date of Birth,First Name,Form of address,Gender,Infix,Last Name,Country,Nationality,Session,Server,Type,`r`n" . this.Report.Value 
                  Path := FileSelect("S 16","Generated Students.csv")
                  try {
                        FileRead(path)
                        FileDelete(path)
                  }
                  FileAppend(export,Path,)
                  Run(path)
                  this.reportWindow.hide()
            }
      }

}



