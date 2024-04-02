#Include utils.ahk

class person extends Object {

      __New(){
            this.systemName := A_Args[1]
            this.session := A_Args[2]
            this.type := A_Args[4]
            this.commandParameter := A_Args[7]
            
            if IsInteger(A_Args[6]) {
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


            generateDateofBirth() {
                  return Format("{:02}",Random(01,28)) "." Format("{:02}",Random(01,12)) "." Random(1980,2005) 
            }

                  
            generatePostCode(){

                  switch this.personalData.nationality {

                  case "GB":
                        return StrUpper(SubStr(this.address.city,1,2)) . Format("{:02}",Random(01,99)) . " " . Random(1,9) . generateLetters(2)
                  
                  case "NL":
                        return Random(1000,9999) . " " . generateLetters(2)

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
}

