require 'rubygems'
require 'watir'


def datatoexcelDBpersonalizations


  @numberoflines = 0

  if @personalizationtypeid > '1'


    excel = WIN32OLE::new("excel.Application")
    wrkbook = excel.Workbooks.Open('G:\SSAutomationPulse.xlsx')
    wrksheet = wrkbook.worksheets(1)
    wrksheet.select

    wrksheet.Cells(@rows, "A").value = @ordnum
    wrksheet.Cells(@rows, "B").value = "Replacement Order"
    wrksheet.Cells(@rows, "C").value = @itemalias
    wrksheet.Cells(@rows, "D").value = @item_description
    wrksheet.Cells(@rows, "E").value = "Monogrammed Items"
    wrksheet.Cells(@rows, "F").value = @quantity
    wrksheet.Cells(@rows, "G").value = @consultantCID
    wrksheet.Cells(@rows, "H").value = "Katie"
    wrksheet.Cells(@rows, "I").value = "test"


    if @personalizationtypeid == "2"

      wrksheet.Cells(@rows, "J").value = @option1

      if @option3 != nil and @textbox2 == nil

        @numberoflines = 1
        wrksheet.Cells(@rows, "M").value = @textbox1

      elsif @option3 != nil and @textbox2 != nil

        @numberoflines = 2
        wrksheet.Cells(@rows, "M").value = @textbox1
        wrksheet.Cells(@rows, "N").value = @textbox2

      end


    elsif @personalizationtypeid == "3"

      wrksheet.Cells(@rows, "J").value = @option3

      if @option5 != nil and @textbox2 == nil

        @numberoflines = 1
        wrksheet.Cells(@rows, "M").value = @textbox1


      elsif @option5 != nil and @textbox2 != nil

        @numberoflines = 2
        wrksheet.Cells(@rows, "M").value = @textbox1
        wrksheet.Cells(@rows, "N").value = @textbox2

      end

      wrksheet.Cells(@rows, "Q").value = "#{@option1}"+"-"+"#{@designstyle}"
      wrksheet.Cells(@rows, "R").value = @option2


    elsif @personalizationtypeid == "5"

      wrksheet.Cells(@rows, "J").value = @option4

      @textbox1 = @option3
      wrksheet.Cells(@rows, "Q").value = @school
      wrksheet.Cells(@rows, "R").value = "White"

    elsif @personalizationtypeid == "14"

      wrksheet.Cells(@rows, "J").value = "White"

      wrksheet.Cells(@rows, "Q").value = "#{@option1}"+"-"+"#{@designstyle}"
      wrksheet.Cells(@rows, "R").value = @option2

    end


    wrksheet.Cells(@rows, "K").value = @embroiderfontstyle
    wrksheet.Cells(@rows, "M").value = @textbox1
    wrksheet.Cells(@rows, "L").value = @numberoflines
    wrksheet.Cells(@rows, "P").value = "0"
    @PulseDesignStyle = wrksheet.Cells(@rows, "Q").value
    @PulseDesignColors = wrksheet.Cells(@rows, "R").value

    wrkbook.save
    excel.quit


  else




  end


end