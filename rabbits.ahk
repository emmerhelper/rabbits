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
            this.tabnames := ["Student","Staff","Options"]
            this.tabs := this.mainWindow.AddTab3("Backgroundffffff",this.tabnames)
            this.tabs.onEvent("Change",resizeWindow)
            this.mainWindow.OnEvent("Close",Quit)

            this.studentTab := rabbitsTab(this,1)
            this.staffTab := rabbitsTab(this,2)

            addOptionsSection()
            
            this.mainWindow.Show()    

            addOptionsSection(){
                  this.tabs.UseTab(3)
                  this.mainWindow.SetFont("w600")
                  this.mainWindow.AddText("ym+24 x+m section","Options")
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
                  if tab.value = 3
                        this.mainWindow.Move(,,230,230)
                  else this.mainWindow.Move(,,746,302)
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
      
                  this.reportWindow.Show("Minimize")
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

            settingsFile(*){
                  RunWait("config.ini")
            }
      }
}



