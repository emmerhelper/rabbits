#Include Person.ahk

class Module extends person {

      Link_To_Staff(){
            this.addRelationship("A","090","P",this.commandParameter)
      }

}