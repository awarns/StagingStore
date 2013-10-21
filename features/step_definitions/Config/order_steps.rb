require 'win32ole'
require 'watir'



def browser

  require 'rubygems'
  require 'watir'

  count = 1

  while count == 1

    @browser= Watir::Browser.new :chrome
    #@browser.speed = :zippy
    @browser.goto("http://testpps.31gifts.corp/admin")


    if @browser.text_field(:id, "ContentPlaceHolder1_txt_username").exists? == false

      @browser.close
      count = 1

    else

      count = 2

    end


  end


end


def login

  @browser.text_field(:id, "ContentPlaceHolder1_txt_username").set("1387")
  @browser.text_field(:id, "ContentPlaceHolder1_txt_password").set("testing31")
  @browser.link(:index, 0).click


end


def retailord

  @browser.goto("https://testpps.31gifts.corp/employee/admin/frm_orderchoose.aspx")
  @browser.link(:text, "Replacement Order").click
  @browser.text_field(:id, "ContentPlaceHolder1_idnobox").set("102091")
  @consultantCID = @browser.text_field(:id, "ContentPlaceHolder1_idnobox").value
  @browser.button(:value, "Enter Order").click

  #@browser.text_field(:id, "ContentPlaceHolder1_ship_fname").set("Test")
  #@browser.text_field(:id, "ContentPlaceHolder1_ship_lname").set("Warns")
  #@browser.text_field(:id, "ContentPlaceHolder1_shipping_info_txt_street_1").set("205 N. nelson rd")
  #@browser.text_field(:id, "ContentPlaceHolder1_shipping_info_txtPostalCode").set("43219")

  while  @browser.link(:id, "ContentPlaceHolder1_btn_save").exists?

    @browser.link(:id, "ContentPlaceHolder1_btn_save").click
    sleep(2)

  end

  browser2 = @browser.frame(:name, "order_top")
  browser2.select_list(:id, "PriceLevelList").select("Replacement Price")


end


def randproduct


  #Excel is linked to Item master Report and is pulling the SKU directly from that excel sheet. The Row
  # in Item master will correlate to the excel in SSAutomation that is wrote to later in the process.

  excel = WIN32OLE::new("excel.Application")
  wrkbook = excel.Workbooks.Open('G:\Summer 2013 Item Master Report.xlsx')
  wrksheet = wrkbook.worksheets(2)
  wrksheet.select


  @itemalias = wrksheet.Cells(@rows, "F").value
  @quantity = "1"
  @print = wrksheet.Cells(@rows, "I").value
  item_description_1 = wrksheet.Cells(@rows, "G").value
  @item_description = item_description_1 + " - " + @print
  @collegiatetrigger = wrksheet.Cells(@rows, "G").value

  if @collegiatetrigger =~ /Collegiate/

    @style = wrksheet.Cells(@rows, "A").value
    browser2 = @browser.frame(:name, "order_top")
    browser2.text_field(:id, "Itemcode").set(@style)
    browser2.text_field(:id, "QuantityList").set(@quantity)
    browser2.button(:value, "Add To Order").click
    sleep(2)

  else

    browser2 = @browser.frame(:name, "order_top")
    browser2.text_field(:id, "Itemcode").set(@itemalias)
    browser2.text_field(:id, "QuantityList").set(@quantity)
    browser2.button(:value, "Add To Order").click
    sleep(2)


  end


  wrkbook.close
  excel.quit


end


def randpersonalizationoption

  #Randomly picking a personalization option. Personalization type is assigned based on what is chosen.

  browser2 = @browser.frame(:name, "order_bottom")
  sleep(1)

  if browser2.select_list(:index, 0).exists?


    personalization = browser2.select_list(:index, 0).options
    length = personalization.length

    index = rand(length)
    @personalization = personalization[index].text

    browser2.select_list(:index, 0).select(@personalization)

    if @personalization == "-- SELECT --"

      @personalization = personalization[2].text
      browser2.select_list(:index, 0).select(@personalization)

    end


    if @personalization =~ /Embroidery/

      @personalizationtypeid = "2"

    elsif @personalization =~ /Icon-It with Text/

      @personalizationtypeid = "3"

    elsif @personalization == "Icon-It w/o Text - add $7"

      @personalizationtypeid = "14"

    elsif @personalization == "None"

      @personalizationtypeid = "1"

    elsif @personalization =~ /Spirit/

      @personalization = @print
      browser2.select_list(:index, 0).select(@personalization)
      @personalizationtypeid = "5"

    end

  else

    @personalizationtypeid = "1"

  end




  if browser2.button(:value, "Add To Order").exists?


    browser2.button(:value, "Add To Order").click

  end


end


def randoption1

  #All rand options below account for any personalization type above.  If an option is present it will choose something
  # Checking for collegiate personalization

  browser2 = @browser.frame(:name, "order_bottom")
  sleep(3)
  if  browser2.select_list(:index, 1).exists? and @personalizationtypeid != "5"

    option = browser2.select_list(:index, 1).options
    length = option.length

    index = rand(length)
    @option1 = option[index].text

    if @option1 == "-- SELECT --"

      @option1 = option[2].text

    end

    browser2.select_list(:index, 1).select(@option1)

  elsif browser2.select_list(:index, 0).exists? and @personalizationtypeid == "5"

    option = browser2.select_list(:index, 0).options
    length = option.length

    index = rand(length)
    @option1 = option[index].text

    browser2.select_list(:index, 0).select(@option1)
    browser2.button(:value, "Add To Order").click
    sleep(2)

  end


end


def randoption2

  browser2 = @browser.frame(:name, "order_bottom")
  sleep(1)
  if  browser2.select_list(:index, 2).exists? and @personalizationtypeid != "5"

    option = browser2.select_list(:index, 2).options
    length = option.length

    index = rand(length)
    @option2 = option[index].text

    if @option2 == "-- SELECT --"

      @option2 = option[1].text

    end

    browser2.select_list(:index, 2).select(@option2)

  elsif browser2.select_list(:index, 0).exists? and @personalizationtypeid == "5"

    option = browser2.select_list(:index, 0).options
    length = option.length

    index = rand(length)
    @option2 = option[index].text

    if @option2 == "-- SELECT --"

      @option2 = option[1].text

    end

    browser2.select_list(:index, 0).select(@option2)


  end


end


def randoption3

  browser2 = @browser.frame(:name, "order_bottom")
  sleep(1)
  if  browser2.select_list(:index, 3).exists? and @personalizationtypeid != "5"

    option = browser2.select_list(:index, 3).options
    length = option.length

    index = rand(length)
    @option3 = option[index].text

    if @option3 == "-- SELECT --"

      @option3 = option[1].text

    end

    browser2.select_list(:index, 3).select(@option3)


  elsif browser2.select_list(:index, 1).exists? and @personalizationtypeid == "5"

    option = browser2.select_list(:index, 1).options
    length = option.length

    index = rand(length)
    @option3 = option[index].text

    if @option3 == "-- SELECT --"

      @option3 = option[1].text

    end

    browser2.select_list(:index, 1).select(@option3)

  end


end


def randoption4


  browser2 = @browser.frame(:name, "order_bottom")
  sleep(1)
  if  browser2.select_list(:index, 4).exists? and @personalizationtypeid != "5"

    option = browser2.select_list(:index, 4).options
    length = option.length

    index = rand(length)
    @option4 = option[index].text

    if @option4 == "-- SELECT --"

      @option4 = option[1].text

    end

    browser2.select_list(:index, 4).select(@option4)


  elsif browser2.select_list(:index, 2).exists? and @personalizationtypeid == "5"

    option = browser2.select_list(:index, 2).options
    length = option.length

    index = rand(length)
    @option4 = option[index].text

    if @option4 == "-- SELECT --"

      @option4 = option[1].text

    end

    browser2.select_list(:index, 2).select(@option4)


  end


end

def randoption5

  browser2 = @browser.frame(:name, "order_bottom")
  sleep(1)
  if  browser2.select_list(:index, 5).exists?

    option = browser2.select_list(:index, 5).options
    length = option.length

    index = rand(length)
    @option5 = option[index].text

    if @option5 == "-- SELECT --"

      @option5 = option[1].text

    end


    browser2.select_list(:index, 5).select(@option5)

  end


end

def textbox1

  browser2 = @browser.frame(:name, "order_bottom")
  sleep(1)
  if browser2.text_field(:index, 0).exists?

    browser2.text_field(:index, 0).set("A")
    @textbox1 = browser2.text_field(:index, 0).value


  end


end


def textbox2

  browser2 = @browser.frame(:name, "order_bottom")
  sleep(1)
  if browser2.text_field(:index, 1).exists?

    browser2.text_field(:index, 1).set("W")
    @textbox2 = browser2.text_field(:index, 1).value


  end


end


def submitord

  #Submiting Order. Nothing has been wrote to excel yet.  However everything needed is captured.




  #Grabbing school name if Collegiate from Screen for use later in dbo.Item and Personalizations
  if @personalizationtypeid == "5"

    @school = @browser.frame(:name, "order_bottom").table(:class, "table_pers_style").row(:index, 0).cell(:index, 0).text


    @school = @school.split /- /
    @school = @school[1]
    @school = @school.split /'/
    @school = @school[0]


  end


  browser2 = @browser.frame(:name, "order_bottom")

  if browser2.button(:value, "Done Personalizing").exists?

    browser2.button(:value, "Done Personalizing").click

  end

  @browser.goto("testpps.31gifts.corp/employee/admin/frm_payment.aspx")

  balance = @browser.table(:class, "normal").text
  arr = balance.split /[\$_]/

  @browser.form(:id, "frm_payment").link(:text, "+ Credit Card Payment").click
  @browser.text_field(:id, "txt_cno").set("4111111111111111")
  @browser.text_field(:id, "txt_cdate").set("1215")
  @browser.text_field(:id, "txt_camount").set(arr[3])
  @browser.button(:value, "Save Payment").click
  @browser.button(:value, "Save Order").wait_until_present
  @browser.button(:value, "Save Order").click
  sleep(2)
  while @browser.text.include?("Thank You") == false

    sleep(2)

  end
  ordernum = @browser.div(:id, "div_submit_out").text
  ordernum = ordernum.split (/# */)
  ordernum = ordernum[1].split /[T_]/
  @ordnum = ordernum[0]


  @browser.close


end






