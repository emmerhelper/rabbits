;Worker v0.2
#SingleInstance Off
#NoTrayIcon
#Include lib/Student.ahk
#Include lib/utils.ahk

;; Arguments:

;; A_Args[1] is the system name that sap is connecting to
; msgbox A_Args[1]

;; A_Args[2] is the session of SAP that this instance controls
; msgbox A_Args[2]

;; A_Args[3] is the count (how many students have already been generated)
; msgbox A_Args[3]

;;; A_Args[4] is the process to be run
; msgbox A_Args[4]

;; A_Args[5] is a function parameter, could be multiple things (nationality or student number)
; msgbox A_Args[5]

;; A_Args[6] is a function parameter, could be multiple things (course to enroll into)
; msgbox A_Args[6]

;; Main flow
loop A_Args[3]{
      rabbits .= "🐇"
}
SetTitleMatchMode(3)
ILive := Gui("-SysMenu","Rabbits are working, please wait warmly 🐇")
ILive.AddText("",rabbits)
ILive.Show("X0 Y0 AutoSize")

if IsInteger(A_Args[5]){
      newboy := student(A_Args[5])
      newboy.%A_Args[4]%()
} else {
      newboy := student()
      newboy.generate(A_Args[5])
      newboy.makeStudentFile()
      
      communicating := true
      while communicating {
            try {
                  WinActivate("Rabbits")
                  EditPaste(newboy.number . "`r`n","Edit2","Rabbits")
                  communicating := false
            }
            catch {
                  sleep 100
            }
      }

      csvWait := true

      while csvWait {

            try {

                  WinActivate("Rabbits Report")
                  EditPaste(newboy.number . "," . newboy.personalData.firstName . "," . newboy.personalData.lastName . "," . newboy.personalData.nationality . " `r`n","Edit1","Rabbits Report")
                  csvWait := false
            }
            catch {
                  sleep 100
            }

      }
}
ExitApp


Esc:: ExitApp

