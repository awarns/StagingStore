require 'win32ole'
require 'watir'

def comparedatadbitem


  #comparing items for dbo.item with items saved in excel previously.

  number = 1

  if number == 1

    @excel = WIN32OLE::new("excel.Application")
    wrkbook = @excel.Workbooks.Open('G:\SSAutomation.xlsx')
    wrksheet = wrkbook.worksheets(2)
    wrksheet.select
  else


  end

  if @ordnum == ""

    @ordnum = 'test'

  end

  open('31StagingStore', 'cmh2wdsql02')
  query("select top 1 OrderSystemUUID, PersonalizationTypeID, ItemAlias, LineNumber, LineSeq, Qty, Barcode, EmbroiderFontColor, EmbroiderFontStyle,EmbroiderLine1, EmbroiderLine2, EmbroiderLine3,IsKitHeader, IsKitComponent, KitHeader, TextExpression, Kid1, Kid1Text, Kid2, Kid2Text, Kid3, Kid3Text, Kid4, Kid4Text, Kid5, Kid5Text, Kid6, Kid6Text, StationaryStyle,ItemStatus, DesignStyle, DesignColor, ParentItemID, KidTextPrefix, Price from [31StagingStore].[dbo].[Item]where Barcode like ('#{@barcode}'+ '%')")
  close
  @data.flatten
  #puts @data

  #puts @data


  @DBOrdersystemid = @data[0].to_s
  @DBPersonalizationtypeid = @data[1].to_i
  @DBItemAlias = @data[2].to_s
  @DBLineNumber = @data[3].to_i
  @DBLineSeq = @data[4].to_i
  @DBQty = @data[5].to_i
  @DBBarcode = @data[6].to_i
  @DBEmbroidFontColor = @data[7]
  @DBEmbroidFontStyle = @data[8]
  @DBEmbroidline1 = @data[9]
  @DBEmbroidline2 = @data[10]
  @DBEmbroidline3 = @data[11]
  @DBIsKitHeader = @data[12].to_s
  @DBIsKitComponent = @data[13].to_s
  @DBKitHeader = @data[14]
  @DBTextExpression = @data[15]
  @DBItemStatus = @data[29].to_s
  @DBDesignStyle = @data[30]
  @DBDesignColor = @data[31]




  errormessage = "The following Error(s) Occured\n"

  if @DBOrdersystemid == ""


    #puts "Order Number " + @ordnum + " Not found in the Database."
    wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
    errormessage << "Order Number " + @OrdersystemUUID + " Not found in the Database.\n"


  else


    if  @OrdersystemUUID == @DBOrdersystemid


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 4

    else

      puts @OrdersystemUUID
      puts @DBOrdersystemid
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Order System UUID is incorect. Database has " + @DBOrdersystemid + "\n"


    end


    if @personalizationtypeid != @DBPersonalizationtypeid

      @personalizationtypeid = @personalizationtypeid.to_s
      @DBPersonalizationtypeid = @DBPersonalizationtypeid.to_s


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Personalization Type ID is incorrect. The Order was placed with pers type " + @personalizationtypeid + " But the DB returned " + @DBPersonalizationtypeid + "\n"


    end


    if @itemalias != @DBItemAlias

      puts @itemalias
      puts @DBItemAlias

      @itemalias = @itemalias.to_s
      @DBItemAlias = @DBItemAlias.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Item Alias is incorrect. Order was placed with Alias " + @itemalias + " But the DB returned " + @DBItemAlias + "\n"


    end


    if @linenumber != @DBLineNumber

      @linenumber = @linenumber.to_s
      @DBLineNumber = @DBLineNumber.to_s


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Line Number is incorrect. Order has line number " + @linenumber + " But the DB returned " + @DBLineNumber + "\n"

    end


    if @lineseq != @DBLineSeq


      @lineseq = @lineseq.to_s
      @DBLineSeq = @DBLineSeq.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Line Seq is incorrect. Order has line seq " + @lineseq + " But the DB returned " + @DBLineSeq + "\n"


    end


    if @qty != @DBQty


      @qty = @qty.to_s
      @DBQty = @DBQty.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Qty is incorrect. Order has qty of " + @qty + " But the DB returned qty of " + @DBQty + "\n"

    end

    if @barcode != @DBBarcode

      @barcode = @barcode.to_s
      @DBBarcode = @DBBarcode.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Barcode is incorrect. Order has barcode of " + @barcode + " But the DB returned a barcode of " + @DBBarcode + "\n"


    end

    if  @DBEmbroidFontColor == @embroidfontcolor

    else

      @embroidfontcolor = @embroidfontcolor.to_s
      @DBEmbroidFontColor = @DBEmbroidFontColor.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBEmbroidFontColor == nil

        errormessage << "Emrboid. Font Color is incorrect. Order has color of " + @embroidfontcolor + " But the DB returned a color of Null\n"

      elsif @DBEmbroidFontColor != nil


        errormessage << "Emrboid. Font Color is incorrect. Order has color of " + @embroidfontcolor + " But the DB returned a color of " + @DBEmbroidFontColor + "\n"

      end

    end


    if  @embroidfontstyle == @DBEmbroidFontStyle

    else

      @embroidfontstyle = @embroidfontstyle.to_s
      @DBEmbroidFontStyle = @DBEmbroidFontStyle.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBEmbroidFontStyle == nil

        errormessage << "Emrboid. Font Style is incorrect. Order has style of " + @embroidfontstyle + " But the DB returned a style of Null\n"

      elsif @DBEmbroidFontStyle != nil


        errormessage << "Emrboid. Font Style is incorrect. Order has style of " + @embroidfontstyle + " But the DB returned a style of " + @DBEmbroidFontStyle + "\n"

      end

    end

    if  @embroidline1 == @DBEmbroidline1

    else

      @embroidline1 = @embroidline1.to_s
      @DBEmbroidline1 = @DBEmbroidline1.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBEmbroidline1 == nil

        errormessage << "Emrboid. Line 1 is incorrect. Order has line of " + @embroidline1 + " But the DB returned a line of Null\n"

      elsif @DBEmbroidline1 != nil

        @DBEmbroidline1 = @DBEmbroidline1.to_s
        errormessage << "Emrboid. Line 1 is incorrect. Order has line of " + @embroidline1 + " But the DB returned a line of " + @DBEmbroidline1 + "\n"

      end

    end

    if  @embroidline2 == @DBEmbroidline2

    else

      @embroidline2 = @embroidline2.to_s
      @DBEmbroidline2 = @DBEmbroidline2.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBEmbroidline2 == nil

        errormessage << "Emrboid. Line 2 is incorrect. Order has line of " + @embroidline2 + " But the DB returned a line of Null\n"

      elsif @DBEmbroidline2 != nil

        @DBEmbroidline2 = @DBEmbroidline2.to_s
        errormessage << "Emrboid. Line 2 is incorrect. Order has line of " + @embroidline2 + " But the DB returned a line of " + @DBEmbroidline2 + "\n"

      end

    end

    if  @embroidline3 == @DBEmbroidline3

    else

      @embroidline3 = @embroidline3.to_s
      @DBEmbroidline3 = @DBEmbroidline3.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBEmbroidline3 == nil

        errormessage << "Emrboid. Line 3 is incorrect. Order has line of " + @embroidline3 + " But the DB returned a line of Null\n"

      elsif @DBEmbroidline3 != nil

        @DBEmbroidline3 = @DBEmbroidline3.to_s
        errormessage << "Emrboid. Line 3 is incorrect. Order has line of " + @embroidline3 + " But the DB returned a line of " + @DBEmbroidline3 + "\n"

      end

    end


    if @iskitheader != @DBIsKitHeader


      @iskitheader = @iskitheader.to_s
      @DBIsKitHeader = @DBIsKitHeader.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Is Kit Header is incorrect. Order has Is Kit Header of " + @iskitheader + " But the DB returned a value of " + @DBIsKitHeader + "\n"


    end


    if @iskitcomponent != @DBIsKitComponent

      @iskitcomponent = @iskitcomponent.to_s
      @DBIsKitComponent = @DBIsKitComponent.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Is Kit Compnent is incorrect. Order has Is Kit Component of " + @iskitcomponent + " But the DB returned a value of " + @DBIsKitComponent + "\n"


    end


    if  @kitheader == @DBKitHeader

    else

      @kitheader = @kitheader.to_s
      puts @kitheader

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBKitHeader == nil

        errormessage << "Kit Header is incorrect. Order has Kit Header of " + @kitheader + " But the DB returned a value of Null\n"

      elsif @DBKitHeader != nil

        @DBKitHeader = @DBKitHeader.to_s
        errormessage << "Kit Header is incorrect. Order has Kit Header of " + @kitheader + " But the DB returned a value of " + @DBKitHeader + "\n"

      end

    end

    if  @textepression == @DBTextExpression

    else

      @textepression = @textepression.to_s


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBTextExpression == nil
        errormessage << "Text Expression is incorrect. Order has Text Expression of " + @textepression + " But the DB returned a value of Null\n"

      elsif @DBTextExpression != nil

        @DBTextExpression = @DBTextExpression.to_s
        errormessage << "Text Expression is incorrect. Order has Text Expression of " + @textepression + " But the DB returned a value of " + @DBTextExpression + "\n"

      end

    end


    if @ItemStatus != @DBItemStatus

      @ItemStatus = @ItemStatus.to_s
      @DBItemStatus = @DBItemStatus.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Item Status is incorrect. Order has Status of " + @ItemStatus + " But the DB returned a value of " + @DBItemStatus + "\n"


    end

    if  @DesignStyle == @DBDesignStyle

    else

      @DesignStyle = @DesignStyle.to_s


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBDesignStyle == nil

        errormessage << "Design Style is Incorrect. Order has style of " + @DesignStyle + " But the DB returned a value of Null\n"

      elsif @DBDesignStyle != nil

        @DBDesignStyle = @DBDesignStyle.to_s
        errormessage << "Design Style is Incorrect. Order has style of " + @DesignStyle + " But the DB returned a value of " + @DBDesignStyle + "\n"

      end

    end


    if  @DesignColor == @DBDesignColor

    else

      @DesignColor = @DesignColor.to_s


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3

      if @DBDesignColor == nil

        errormessage << "Design Color is Incorrect. Order has color of " + @DesignColor + " But the DB returned a value of Null\n"

      elsif @DBDesignColor != nil

        @DBDesignColor = @DBDesignColor.to_s
        errormessage << "Design Color is Incorrect. Order has color of " + @DesignColor + " But the DB returned a value of " + @DBDesignColor + "\n"

      end

    end


  end


  wrksheet.Cells("#{@rows}", "AK").value = errormessage

  wrkbook.save
  sleep(1)
  @excel.quit
  @excel.visible = 0




  number = number + 1


end