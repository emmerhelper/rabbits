#Include utils.ahk

class person extends Object {

      __New(){
            this.systemName := A_Args[1]
            this.session := A_Args[2]
            this.type := A_Args[4]

            switch this.type {
                  case "Student":
                        this.objectType := "ST"
                  case "Staff":
                        this.objectType := "P"
                  case "Module":
                        this.objectType := "SM"
            }
      
            
            if (A_Args.Length >= 7)
                  this.commandParameter := A_Args[7]
            
            if (A_Args[6]) {
                  this.number := A_Args[6]
            }
            else this.number := false 
      }


      generate(){

            generatePersonalData()
            generateAddress()      

            generatePersonalData(){

                  this.personalData := {}
                  
                  this.personalData.gender := Random(1,2)
                  this.personalData.nationality := A_Args[6]
                  generateLastName()
                  this.personalData.firstName := generateFirstName()
                  this.personalData.dateOFBirth := generateDateofBirth()
                  this.personalData.birthplace := lineFromFile(this.personalData.nationality,"city")
                  this.personalData.nationalityCode := this.personalData.nationality
                  this.personalData.formOfAddress := generateFormOfAddress()
            }
      

            generateAddress() {

                  this.address := {}

                  this.address.HouseNumber := Random(1,149)
                  this.address.Street := lineFromFile(this.personalData.nationality,"street")
                  this.address.City := lineFromFile(this.personalData.nationality,"City")
                  this.address.PostCode := generatePostCode()
                  this.address.PhoneNumber := "06" . Random(10000000,99999999)
                  this.address.Email := generateEmail()
                  

            }

            generateFirstName(){

                  if (this.personalData.gender = 1)
                        return  lineFromFile(this.personalData.nationality,"firstnamemale")
                  else return lineFromFile(this.personalData.nationality,"firstnamefemale")
      
            }

            generateLastName(){

                  lastName := lineFromFile(this.personalData.nationality,"lastName")
                  this.personalData.infix := ""

                  if !InStr(lastName," "){
                        this.personalData.lastName := StrTitle(lastName)
                        return 
                  } else {
                        lastNameArray := StrSplit(lastName," ")
                        this.personalData.lastName := StrTitle(lastNameArray.Pop())
                        for k, v in lastNameArray {
                              if A_Index > 1
                                    this.personalData.infix .= " "

                              this.personalData.infix .= StrUpper(v)
                        }
                  }

            }

            generateFormOfAddress(){
                  if (this.personalData.gender = 1)
                        return "Mr."
                  else return "Mrs."
            }


            generateDateofBirth() {
                  return Format("{:02}",Random(01,28)) "." Format("{:02}",Random(01,12)) "." Random(2005,2007) 
            }

                  
            generatePostCode(){

                  switch this.personalData.nationality {

                  case "GB":
                        return StrUpper(SubStr(this.address.city,1,2)) . Format("{:02}",Random(01,99)) . " " . Random(1,9) . generateLetters(2)
                  
                  case "NL":
                        return Random(1000,9999) . " " . generateLetters(2)
                  
                  case "CA":
                        return generateLetters(1) . Random(0,9) . generateLetters(1) . " " . Random(0,9) . generateLetters(1) . Random(0,9)
           
                  Default:
                        return Random(10000,99999)
                  }
            }

            generateEmail(){
                  email .= this.personalData.firstName
                  email .= "."
                  email .= this.personalData.lastName
                  email.= "@campus.ntt.com"
                  return StrLower(email)
            }

      }

      phoneHome(){
            this.sendNumber()
            this.sendReport()
      }
  
      sendNumber(){
           
            if this.type = "Student"
                  target := "Edit2"
            else target := "Edit4"

            communicating := true
            while communicating {
                  try {
                        EditPaste(this.number . "`r`n",target,"Rabbits")
                        communicating := false
                  }
                  catch {
                        sleep 100
                  }
            }
      }

      sendReport(){
            csvWait := true
            while csvWait {
                        EditPaste(this.getReportString(),"Edit1","Rabbits Report")
                        csvWait := false
                        sleep 100
            }
      }

      getReportString(){
            report := ""
            for k, v in this.OwnProps(){
                  if (IsObject(v)){
                        for i, j in v.OwnProps(){
                              report .= j ","
                        }
                  } else {
                        report .= v ","
                  }
            }
            report .= " `r`n"

            return report
      }

      addRelationship(relationshipType,relationship,relatedObjectType,relatedObjectID){

            session := sapConnect(this.systemName,this.session)
            session.startTransaction("PP02")
            userArea := session.findByID("wnd[0]/usr")

            findTextElement(userArea,'PPHDR-OTYPE').text := this.objectType

            findTextElement(userArea,'PM0D1-SEARK').SetFocus()
            session.findById("wnd[0]").sendVKey(4)
            session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB003").select()
            session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB003/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").text := this.number

            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)

            findTextElement(userArea,'PPHDR-ISTAT').text := "1"
            findTextElement(userArea,'PPHDR-INFTY').text := "1001"
            findTextElement(userArea,'PPHDR-BEGDA').text := IniRead("config.ini","Relationships","valid_from")

            session.findByID("wnd[0]/tbar[1]/btn[5]").press()

            findTextElement(userArea,'P1001-RSIGN').text := relationshipType
            findTextElement(userArea,'P1001-RELAT').text := relationship
            userArea.FindByName('P1001-SCLAS','GuiComboBox').Key := relatedObjectType
            findTextElement(userArea,'P1001-SOBID').text := relatedObjectID

            session.findById("wnd[0]").sendVKey(11)
            try session.findById("wnd[0]").sendVKey(11)



      }
}

