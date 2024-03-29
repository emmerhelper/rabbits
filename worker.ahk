;; Worker v1.0
;; Multi-threading in AHK? That's right!
#SingleInstance Off
#NoTrayIcon
#Include lib/Student.ahk
#Include lib/utils.ahk
#Include lib/Staff.ahk


;; Arguments:
;; A_Args[1] is the system name that sap is connecting to
;; A_Args[2] is the session of SAP that this instance controls
;; A_Args[3] is the count (how many have already been generated)
;; A_Args[4] is the type to generate
;; A_Args[5] is the process to be run
;; A_Args[6] is a command parameter, could be multiple things (nationality or student number)
;; A_Args[7] is a command parameter, - the value of the command
;; Check params with following function:
showAllParams(A_Args)
;; -----------------------------------------

;; Main flow ------------------------------
drawCountGUI()
newboy := %A_Args[4]%()
newboy.%A_Args[5]%()
ExitApp()
;; ----------------------------------------

drawCountGUI(){
      loop A_Args[3]{
            rabbits .= "🐇"
      }
      SetTitleMatchMode(3)
      ILive := Gui("-SysMenu","Rabbits are working, please wait warmly 🐇")
      ILive.AddText("",rabbits)
      ILive.Show("X0 Y0 AutoSize")
}

;; Killswitch
Esc:: ExitApp

