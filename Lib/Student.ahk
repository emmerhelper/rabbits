#Include Person.ahk

class student extends person {

      create(){
            this.generate()
            this.makeStudentFile()
            this.phoneHome()
      }

      makeStudentFile(){
            ;; Open the transaction
            session := sapConnect(this.systemName,this.session)
            session.startTransaction("PIQSTC")

            userArea := session.findByID("wnd[0]/usr")

            ;; Fill out all necessary fields:

            ;; Personal data tab
            userArea.findByName("P1702-GESCH","GuiComboBox").Key := this.personalData.gender
            userArea.findByName("P1702-NACHN","GuiTextField").Text := this.personalData.lastName
            userArea.findByName("P1702-VORNA","GuiTextField").Text := this.personalData.FirstName
            userArea.findByName("P1702-RUFNM","GuiTextField").Text := this.personalData.FirstName
            userArea.findByName("P1702-GBDAT","GuiCTextField").Text := this.personalData.DateOfBirth
            userArea.findByName("P1702-NATIO","GuiComboBox").Key := this.personalData.NationalityCode
            userArea.findByName("P1702-GBORT","GuiTextField").Text := this.personalData.birthplace

            if this.personalData.infix {
                  userArea.findByName("P1702-VORSW","GuiComboBox").Key := this.personalData.infix
            }

            ;; Standard Address Tab

            userArea.findByName("DETLTAB02","GuiTab").Select()
            findTextElement(userArea,"ADDR2_DATA-STREET").Text := this.Address.Street
            findTextElement(userArea,"ADDR2_DATA-HOUSE_NUM1").Text := this.Address.HouseNumber
            findTextElement(userArea,"ADDR2_DATA-POST_CODE1").Text := this.Address.PostCode
            findTextElement(userArea,"ADDR2_DATA-CITY1").Text := this.Address.City
            findTextElement(userArea,"ADDR2_DATA-COUNTRY").Text := this.personaldata.Nationality
            findTextElement(userArea,"SZA11_0100-MOB_NUMBER").Text := this.address.PhoneNumber
            findTextElement(userArea,"SZA11_0100-SMTP_ADDR").Text := this.address.Email
      
            ;; Save
            session.findByID("wnd[0]/tbar[0]/btn[11]").press()
            
            while !this.number{
                  sleep 100
                  try this.number := session.findByID("wnd[0]/usr/subMAINSCREEN:SAPLHRPIQ00STUDENT_NF_MD:2000/subWORKSPACE:SAPLHRPIQ00STUDENT_NF_MD:2200/subOVERVIEW:SAPLHRPIQ00STUDENT_NF_MD:1001/txtPIQ1000-STUDENT12").text 
            }

      }

      Set_Home_Student(){

            MultipleParameters := StrSplit(this.commandParameter,",")
            this.residentCountry := MultipleParameters[1]
            this.residentState := MultipleParameters[2]

            session := sapConnect(this.systemName,this.session)
            session.startTransaction("PIQSTM")
            userArea := session.findByID("wnd[0]/usr")
            findTextElement(userArea,"PIQ1000-STUDENT12").SetFocus()
            findTextElement(userArea,"PIQ1000-STUDENT12").Text := this.number
            session.findById("wnd[0]").sendVKey(0)
            selectStudentFileTab(userArea,"Visa")
            userArea.findByName("P1711-RESID_COUNTRY","GuiComboBox").Key := this.residentCountry
            findTextElement(userArea,"P1711-RESID_REGIO").Text := this.residentState
            session.findByID("wnd[0]/tbar[0]/btn[11]").press()

      }

      ZPIQSU01(){

            session := sapConnect(this.systemName,this.session)
            session.startTransaction("ZPIQSU01")
            userArea := session.findByID("wnd[0]/usr")
            findTextElement(userArea,"P_STNUM").Text := this.number
            session.findById("wnd[0]").sendVKey(11)

      }

      
      Register(){
            ini := returnIniSectionAsObject("config","Register")

            session := sapConnect(this.systemName,this.session)
            userArea := openStudentInStudentFile(session, this.number)

            selectStudentFileTab(userArea,"Admission")
            userArea.findByName("CONTAINER_ADM_LIST","GuiCustomControl").children[0].children[0].pressToolbarContextButton("PB_ADM_CREATE")
            userArea.findByName("CONTAINER_ADM_LIST","GuiCustomControl").children[0].children[0].selectContextMenuItem("PB_ADM_APPLY")
            userArea1 := session.findByID("wnd[1]/usr")
            findTextElement(userArea1,"PIQSTREGDIAL-SC_SHORT").Text := this.commandParameter
            session.findById("wnd[1]").sendVKey(0)

            ;; Set registration year, period, type and category, otherwise study routes won't work
            try {
                  userArea1.findByName("PIQSTADM-ADM_AYEAR","GuiComboBox").Key := ini.academic_year
            } catch {
            userArea1.findByName("PIQSTADM-ADM_AYEAR","GuiComboBox").Key := userArea1.findByName("PIQSTADM-ADM_AYEAR","GuiComboBox").Entries[0].Key
            }

            try {
                  userArea1.findByName("PIQSTADM-ADM_PERID","GuiComboBox").Key := ini.academic_session
            } catch {
                  userArea1.findByName("PIQSTADM-ADM_PERID","GuiComboBox").Key := userArea1.findByName("PIQSTADM-ADM_PERID","GuiComboBox").Entries[0].Key
            }

            userArea1.findByName("PIQSTADM-ADM_ENRCATEG","GuiComboBox").Key := "01"
            userArea1.findByName("PIQSTADM-ADM_CATEG","GuiComboBox").Key := "01"



            session.findById("wnd[1]/tbar[0]/btn[11]").press()
      }

      Admit(){

            session := sapConnect(this.systemName,this.session)
            userArea := openStudentInStudentFile(session, this.number)

            selectStudentFileTab(userArea,"Admission")
            admissionsTable := userArea.findByName("CONTAINER_ADM_LIST","GuiCustomControl").children[0].children[0]

            ;; Click the correct row of the table
            loop admissionsTable.rowCount{
                  if (admissionsTable.getCellValue(A_Index-1,"SC_SHORT") = this.commandParameter){
                        admissionsTable.currentCellRow := A_Index-1
                  }
            }

            admissionsTable.pressToolbarButton("PB_ADM_DETAIL")






            session.findById("wnd[0]/usr/subAUDIT_PROFILE_DATA:SAPLHRPIQ00AUDITFORMS_PROFDIAL:0100/cntlC_CONT_PROFILE/shellcont/shell/shellcont[1]/shell[0]").pressButton("FC_GREEN")
            while (session.findByID("wnd[0]/sbar").text != "Data was saved"){
                  session.findById("wnd[0]").sendVKey(11)
                  sleep 100
            }
            session.findById("wnd[0]/tbar[1]/btn[13]").press()
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

      }





}