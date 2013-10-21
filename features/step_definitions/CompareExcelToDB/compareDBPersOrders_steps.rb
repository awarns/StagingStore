require 'rubygems'
require 'watir'


def compareDBpersOrders


  number = 1

  if number == 1

    @excel = WIN32OLE::new("excel.Application")
    wrkbook = @excel.Workbooks.Open('G:\SSAutomationPulse.xlsx')
    wrksheet = wrkbook.worksheets(2)
    wrksheet.select

  else


  end


  puts @OrderNumber
  open('pulse', 'cmh2wdpulse01\pulse')
  query("select Product, OrderNumber, Job, style, ShortStyle, Hoop, TextLine1, TextLine2, Thread, Status, DesignStyle, DesignColors from pulse.dbo.Orders where OrderNumber = '#{@OrderNumber}'")
  close
  @data.flatten

  puts @data



  @DBproduct = @data[0]
  @DBOrderNumber = @data[1].to_s
  @DBJob = @data[2]
  @DBStyle = @data[3]
  @DBShortStyle = @data[4]
  @DBHoop = @data[5]
  @DBTextLine1 = @data[6]
  @DBTextLine2 = @data[7]
  @DBThread = @data[8]
  @DBStatus = @data[9]
  @DBDesignStyle = @data[10]
  @DBDesignColors = @data[11]



  if @DBTextLine1 == ""

    @DBTextLine1 = nil

  end




  if @DBTextLine2 == ""

    @DBTextLine2 = nil

  end



  errormessage = "The following Error(s) Occured\n"


  if @DBOrderNumber == ""



    wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
    errormessage << "Order Number " + @OrderNumber + " Not found in the Database.\n"


  else


    if @OrderNumber == @DBOrderNumber


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 4



    end

    if @product != @DBproduct


      @product = @product.to_s
      @DBproduct = @DBproduct.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Product is incorrect. Order was placed with #{@product}, but the db returned #{@DBproduct}.\n"


    end

    if @DBJob != @job
      
      @DBJob = @DBJob.to_s
      @job = @job.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Job is incorrect. Order has job number #{@job}, but the db returned #{@DBJob}.\n"
      
      
    end

    if @DBStyle != @style

      @DBStyle = @DBStyle.to_s
      @style = @style.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Style is incorrect. Order has style of #{@style}, but the db returned #{@DBStyle}.\n"


    end

    if @DBShortStyle != @shortstyle

      @DBShortStyle = @DBShortStyle.to_s
      @shortstyle = @shortstyle.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Short style is incorrect. Order has style #{@shortstyle}, but the db returned #{@DBShortStyle}.\n"

    end


    if @DBHoop != @hoop

      @DBHoop = @DBHoop.to_s
      @hoop = @hoop.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Hoop is incorrect. Order has hoop of #{@hoop}, but the db returned #{@DBHoop}.\n"


    end

    if @DBTextLine1 != @textline1

      @DBTextLine1 = @DBTextLine1.to_s
      @textline1 = @textline1.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Text Line 1 is incorrect. Order has text of #{@textline1}, but the db returned #{@DBTextLine1}.\n"


    end

    if @DBTextLine2 != @textline2

      @DBTextLine2 = @DBTextLine2.to_s
      @textline2 = @DBTextLine2.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Text Line 2 is incorrect. Order has text of #{@textline2}, but the db returned #{@DBTextLine2}.\n"

    end

    if @DBThread != @thread

      @DBThread = @DBThread.to_s
      @thread = @thread.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Thread is incorrect. Order has thread of #{@thread}, but the db returned #{@DBThread}.\n"


    end

    if @DBStatus != @status

      @DBStatus = @DBStatus.to_s
      @status = @status.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Status is incorrect. Order has status of #{@status}, but the db returned #{@DBStatus}.\n"


    end

    if @DBDesignStyle != @designStyle

      @DBDesignStyle = @DBDesignStyle.to_s
      @designStyle = @designStyle.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Design style is incorrect. Order has style of #{@designStyle}, but the db returned #{@DBDesignStyle}"



    end

    if @DBDesignColors != @designColors

      @DBDesignColors =@DBDesignColors.to_s
      @designColors = @designColors.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Design Color is incorrect. Order has color of #{@designColors}, but the DB returned #{@DBDesignColors}"

    end


  end
  wrksheet.Cells("#{@rows}", "N").value = errormessage

  wrkbook.save
  sleep(1)
  @excel.quit()



  number = number + 1


end