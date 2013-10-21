require 'win32ole'
require 'watir'



def compareDBPers

  number = 1

  if number == 1

    @excel = WIN32OLE::new("excel.Application")
    wrkbook = @excel.Workbooks.Open('G:\SSAutomationPulse.xlsx')
    wrksheet = wrkbook.worksheets(1)
    wrksheet.select

  else


  end
  open('ThirtyOne_Personalization', 'cmh2wdpulse01\pulse')
  query("select order_number, order_type, item_number, item_description, product_type, qty, consultant_id, first_name, last_name, font_color, font_style, number_of_lines,text_line_1, text_line_2, embroider_dt_tm, FileDeleted, DesignStyle, DesignColors from [ThirtyOne_Personalization].[dbo].[personalizations] where order_number in ('#{@ordnum}')")
  close
  @data.flatten


  puts @data


  @DBorder_number = @data[0]
  @DBorder_type = @data[1]
  @DBitem_number = @data[2]
  @DBitem_description = @data[3]
  @DBproduct_type = @data[4]
  @DBqty = @data[5].to_i
  @DBconsultant_id = @data[6]
  @DBfirst_name = @data[7]
  @DBlast_name = @data[8]
  @DBfont_color = @data[9]
  @DBfont_style = @data[10]
  @DBnumber_of_lines = @data[11].to_i
  @DBtext_line_1 = @data[12]

  if @DBtext_line_1 == ""

    @DBtext_line_1 = nil

  end



  @DBtext_line_2 = @data[13]

  if @DBtext_line_2 == ""

    @DBtext_line_2 = nil

  end


  @DBembroid_dt_tm = @data[14]
  @DBFileDeleted = @data[15].to_s
  @DBDesignStyle = @data[16]
  @DBDesignColors = @data[17]



  errormessage = "The following Error(s) Occured\n"


  if @DBorder_number == ""



    wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
    errormessage << "Order Number " + @ordnum + " Not found in the Database.\n"


  else


    if @ordnum == @DBorder_number


      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 4



    end


    if @DBorder_type != @order_type

      @DBorder_type = @DBorder_type.to_s
      @order_type = @order_type.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Order_type is incorrect. Order was placed as a #{@order_type}, but the db returned #{@DBorder_type}.\n"



    end

    if @DBitem_number != @item_number

      @DBitem_number = @DBitem_number.to_s
      @item_number = @item_number.to_s

      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Item_number is incorrect. Order was placed with item number #{@item_number}, but the db returned #{@DBitem_number}.\n"


    end


    if @DBitem_description != @item_description

      @DBitem_description = @DBitem_description.to_s
      @item_description = @item_description.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Item_description is incorrect. Order was placed with item description #{@item_description}, but the db returned #{@DBitem_description}.\n"


    end


    if @DBproduct_type != @product_type

      @DBproduct_type = @DBproduct_type.to_s
      @product_type = @product_type.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Product_type is incorrect.  Order was placed with type #{@product_type}, but the db returned #{@DBproduct_type}.\n"


    end

    if @DBqty != @qty

      @DBqty = @DBqty.to_s
      @qty = @qty.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Qty is incorrect. Order was placed with #{@qty}, but the db returned #{@DBqty}.\n"


    end

    if @DBconsultant_id != @consultant_id

      @DBconsultant_id = @DBconsultant_id.to_s
      @consultant_id = @consultant_id.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Consultant_ID is incorrect. Order was placed by #{@consultant_id}, but the db returned #{@DBconsultant_id}/\n"

    end

    if @DBfirst_name != @first_name

      @DBfirst_name = @DBfirst_name.to_s
      @first_name = @first_name.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "First name is incorrect. Order was placed by #{@first_name}, but the db returned #{@DBfirst_name}.\n"

    end

    if @DBlast_name != @last_name

      @DBconsultant_id = @DBconsultant_id.to_s
      @last_name = @last_name.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Last name is incorrect. Order was placed by #{@last_name}, but the db returned #{@DBlast_name}.\n"

    end

    if @DBfont_color != @font_color

      @DBfont_color = @DBfont_color.to_s
      @font_color = @font_color.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Font_Color is incorrect.  Order was placed with #{@font_color}, but the db returned #{@DBfont_color}.\n"

    end

    if @DBfont_style != @font_style

      @DBfont_style = @DBfont_style.to_s
      @font_style = @font_style.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Font_Style is incorrect. Order was placed with font style #{@font_style}, but the db returned #{@DBfont_style}.\n"
    end

    if @DBnumber_of_lines != @number_of_lines

      @DBnumber_of_lines = @DBnumber_of_lines.to_s
      @number_of_lines = @number_of_lines.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Number_of_lines is incorrect. Order has #{@number_of_lines}, but the db returned #{@DBnumber_of_lines}.\n"
    end

    if @DBtext_line_1 != @text_line_1

      @DBtext_line_1 = @DBtext_line_1.to_s
      @text_line_1 = @text_line_1.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Text_line_1 is incorrect. Order has text of #{@text_line_1}, but the db returned #{@DBtext_line_1}.\n"

    end

    if @DBtext_line_2 != @text_line_2

      @DBtext_line_2 = @DBtext_line_2.to_s
      @text_line_2 = @text_line_2.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Text_line_2 is incorrect. Order has text of#{@text_line_2}, but the db returned #{@DBtext_line_2}.\n"

    end

    if @DBembroid_dt_tm != @embroid_dt_tm

      @DBembroid_dt_tm = @DBembroid_dt_tm.to_s
      @embroid_dt_tm = @embroid_dt_tm.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Emrboid_dt_tm is incorect. Order has #{@embroid_dt_tm}, but the db returned a value of #{@DBembroid_dt_tm}.\n"

    end

    if @DBFileDeleted != @fileDeleted

      @DBFileDeleted = @DBFileDeleted.to_s
      @fileDeleted = @fileDeleted.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "File Deleted is incorrect. Order has #{@fileDeleted}, but the db returned a value of #{@DBFileDeleted}.\n"

    end

    if @DBDesignStyle != @DesignStyle

      @DBDesignStyle = @DBDesignStyle.to_s
      @DesignStyle = @DesignStyle.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Design Style is incorrect. Order has #{@DesignStyle}, but the db returned a value of #{@DBDesignStyle}.\n"


    end

    if @DBDesignColors != @DesignColors

      @DBDesignColors = @DBDesignColors.to_s
      @DesignColors = @DesignColors.to_s
      wrksheet.Rows("#{@rows}").Interior.ColorIndex = 3
      errormessage << "Design Colors is incorrect. Order has #{@DesignColors}, but the db returned a vlue of #{@DBDesignColors}.\n"

    end


  end


  wrksheet.Cells("#{@rows}", "S").value = errormessage

  wrkbook.save
  sleep(1)
  @excel.quit()



  number = number + 1


end