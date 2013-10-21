require 'win32ole'
require 'watir'



def extractdataDBPers

  @rows = @rows.to_s
  range1 = "A"+@rows
  range2 = "R"+@rows


  excel = WIN32OLE::new("excel.Application")
  wrkbook = excel.Workbooks.Open('G:\SSAutomationPulse.xlsx')
  wrksheet = wrkbook.worksheets(1)
  wrksheet.select
  @arr = Hash.new
  @arr = wrksheet.Range("#{range1}:#{range2}").value
  wrkbook.close
  excel.quit
  @arr = @arr.flatten


  @ordnum = @arr[0].to_s
  @ordnum = @ordnum.chomp
  @order_type = @arr[1]
  @item_number = @arr[2]
  @item_description = @arr[3]
  @product_type = @arr[4]
  @qty = @arr[5].to_i
  @consultant_id = @arr[6]
  @first_name = @arr[7]
  @last_name = @arr[8]
  @font_color = @arr[9]
  @font_style = @arr[10]
  @number_of_lines = @arr[11].to_i
  @text_line_1 = @arr[12]
  @text_line_2 = @arr[13]
  @embroid_dt_tm = @arr[14]
  @fileDeleted = @arr[15].to_s
  @DesignStyle = @arr[16]
  @DesignColors = @arr[17]




  @rows = @rows.to_i

end
