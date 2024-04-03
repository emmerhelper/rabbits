class staff extends person {
        
      create(){
            this.generate()
            this.addNewP()
            this.phoneHome()
      }

      addNewP(){
            ini := returnIniSectionAsObject("config","New_Staff")

            session := sapConnect(this.systemName,this.session)
            session.startTransaction("PA30")   
            userArea := session.findByID("wnd[0]/usr")
            
            ;; Maintain HR Master Data
            findTextElement(userArea,"RP50G-PERNR").text := ""
            userArea.findByName("SAPMP50ATC_MENU","GuiTableControl").getAbsoluteRow(0).Selected := True
            session.findById("wnd[0]/tbar[1]/btn[5]").press()

            ;; Actions screen
            findTextElement(userArea,"P0000-BEGDA").text := Ini.begin_Date
            findTextElement(userArea,"P0000-MASSG").text := ini.Reason_For_Action
            findTextElement(userArea,"PSPAR-PLANS").text := ini.Position
            findTextElement(userArea,"PSPAR-WERKS").text := ini.Personnel_Area
            findTextElement(userArea,"PSPAR-PERSG").text := ini.Employee_Group
            findTextElement(userArea,"PSPAR-PERSK").text := ini.Employee_Subgroup
            session.findByID("wnd[0]/mbar/menu[0]/menu[1]").Select()
            
            ;; Create Personal Data
            findTextElement(userArea,"Q0002AE-VORNA").text := this.personalData.FirstName
            findTextElement(userArea,"Q0002AE-NACHN").text := this.personalData.lastName
            userArea.findByName("Q0002-ANREX","GuiComboBox").Key := this.personalData.formOfAddress
            findTextElement(userArea,"P0002-GBDAT").text := this.personalData.dateOFBirth
            ; userArea.findByName("P0002-SPRSL","GuiComboBox").Key := this.personalData.nationality
            userArea.findByName("Q0002AE-GESCH","GuiComboBox").Key := this.personalData.gender
            userArea.findByName("P0002-NATIO","GuiComboBox").Key := this.personalData.nationality
            userArea.findByName("Q0002-KITXT","GuiComboBox").Key := ini.Religion
            findTextElement(userArea,"P3255-VISPR").text := ini.profession
            session.findByID("wnd[0]/mbar/menu[0]/menu[1]").Select()
            
            ;; Create organizational assignment
            findTextElement(userArea,"P0001-BTRTL").text := ini.Subarea
            findTextElement(userArea,"P0001-GSBER").text := ini.Business_Area
            findTextElement(userArea,"P0001-ABKRS").text := ini.Payroll_Area
            session.findByID("wnd[0]/mbar/menu[0]/menu[1]").Select()
            
            this.number := findTextElement(userArea,"RP50G-PERNR").text

      }


}