require 'win32ole'
require 'watir'


def comparedatadborder


  #comparing items in dbo.Order to Items saved in excel for dbo.order. Reporting any errors.

  number = 1

  if number == 1

    @excel = WIN32OLE::new("excel.Application")
    wrkbook = @excel.Workbooks.Open('G:\SSAutomation.xlsx')
    wrksheet = wrkbook.worksheets(1)
    wrksheet.select

  else


  end


  open('31StagingStore', 'cmh2wdsql02')
  query("select ordernumber, ConsultantCID, OrderStatus, OrderTypeID from [31StagingStore].[dbo].[Order]where OrderNumber in ('#{@ordnum}')")
  close
  @data.flatten
  puts @data





  errormessage = "The Following Error(s) Occured\n"

  @DBOrdnum = @data[0].to_s
  @DBOrdnum = @DBOrdnum.chomp
  @DBcid = @data[1].to_i
  @DBOrderStatus = @data[2]
  @DBOrderTypeID = @data[3].to_i



  if @DBOrdnum == ""



    wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
    errormessage << "Order Number " + @ordnum + " Not found in the Database.\n"


  else


    if @ordnum == @DBOrdnum


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 4


    else

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << @ordnum + " does not mach. Found" + " " + @DBOrdnum + " In the Database\n"


    end

    if @consultantCID == @DBcid

    else

      @consultantCID = @consultantCID.to_s
      @DBcid = @DBcid.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Error for Order Number " + @ordnum + ". Order was placed with Consultant Id " + @consultantCID + " But ID " + @DBcid + " Was found in the DB\n"


    end

    if @OrderStatus == @DBOrderStatus

    else

      @consultantCID = @consultantCID.to_s
      @DBcid = @DBcid.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Error for Order Number " + @ordnum + " For Column Order Status. Database has a status of " + @DBOrderStatus +"\n"


    end


    if @OrderTypeID == @DBOrderTypeID

    else

      @consultantCID = @consultantCID.to_s
      @DBcid = @DBcid.to_s
      @DBOrderTypeID = @DBOrderTypeID.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Error for Order Number " + @ordnum + " For Column Order Type ID. Database has a typeID of " + @DBOrderTypeID



    end


  end

  wrksheet.Cells("#{@rows}", "F").value = errormessage

  wrkbook.save
  sleep(2)
  @excel.quit
  @excel.visible = 0


  number = number + 1

end