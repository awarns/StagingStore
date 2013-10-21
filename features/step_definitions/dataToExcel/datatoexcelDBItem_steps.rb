require 'rubygems'
require 'watir'

def datatoexcelDBItem


  #formating items captured during order placement and sending them to excel for table dbo.Item.
  #Conditional statements below based on Personalization Type that was assigned earlier.


  excel = WIN32OLE::new("excel.Application")
  wrkbook = excel.Workbooks.Open('G:\SSAutomation.xlsx')
  wrksheet = wrkbook.worksheets(2)
  wrksheet.select


  if @personalizationtypeid != "5"

    @OrderSystemUUID = "#{@ordnum.rstrip}" + "|1|0"
    @Barcode = "#{@ordnum.rstrip}" + "100"


    wrksheet.Cells(@rows, "A").value = @OrderSystemUUID
    wrksheet.Cells(@rows, "B").value = @personalizationtypeid
    wrksheet.Cells(@rows, "C").value = @itemalias
    wrksheet.Cells(@rows, "D").value = "1"
    wrksheet.Cells(@rows, "E").value = "0"
    wrksheet.Cells(@rows, "F").value = @quantity
    wrksheet.Cells(@rows, "G").value = @Barcode


  end


  if @personalizationtypeid == "1"

    wrksheet.Cells(@rows, "M").value = "false"
    wrksheet.Cells(@rows, "N").value = "false"
    wrksheet.Cells(@rows, "AD").value = "Entered"


  end


  if @personalizationtypeid == "2"

    #Query below grabs style out from Xlat table and uses the data returned for font style


    open('31StagingStore', 'cmh2wdsql02')
    query("select style_out from [31StagingStore].[dbo].[FontXlat]where sku = ('#{@itemalias}') and style_in = '#{@option2.rstrip}'")
    close
    @data.flatten


    wrksheet.Cells(@rows, "H").value = @option1
    @embroiderfontstyle = @data[0].to_s
    wrksheet.Cells(@rows, "I").value = @embroiderfontstyle
    wrksheet.Cells(@rows, "J").value = @textbox1.rstrip


    if @textbox2 == nil


    else

      wrksheet.Cells(@rows, "K").value = @textbox2


    end


    wrksheet.Cells(@rows, "M").value = "false"
    wrksheet.Cells(@rows, "N").value = "false"
    wrksheet.Cells(@rows, "AD").value = "Entered"

    #wrksheet.Cells(@rows, "AI").value = "NULL"


  end


  if @personalizationtypeid == "3"


    #Query below grabs style out from Xlat table and uses the data returned for font style

    open('31StagingStore', 'cmh2wdsql02')
    query("select style_out from [31StagingStore].[dbo].[FontXlat]where sku = ('#{@itemalias}') and style_in = '#{@option4.rstrip}'")
    close
    @data.flatten


    wrksheet.Cells(@rows, "H").value = @option3
    @embroiderfontstyle = @data[0].to_s
    wrksheet.Cells(@rows, "I").value = @embroiderfontstyle
    wrksheet.Cells(@rows, "J").value = @textbox1.rstrip


    if @textbox2 == nil


    else

      wrksheet.Cells(@rows, "K").value = @textbox2


    end

    #splitting everything after '-' in font style to add to Design style

    @designstyle = @embroiderfontstyle
    @designstyle = @designstyle.split (/- */)

    wrksheet.Cells(@rows, "M").value = "false"

    @designstyle = @designstyle[1].to_s


    wrksheet.Cells(@rows, "N").value = "false"
    wrksheet.Cells(@rows, "AD").value = "Entered"
    wrksheet.Cells(@rows, "AE").value = "#{@option1}"+"-"+"#{@designstyle}"
    wrksheet.Cells(@rows, "AF").value = @option2
    #wrksheet.Cells(@rows, "AI").value = "NULL"


  end


  if @personalizationtypeid == "5"


    #collegiate


    wrksheet = wrkbook.worksheets(3)
    wrksheet.select

    uuidlastdigit = 0
    barcodelastnum = 100
    lineseq = 0

    tempitemalias = @itemalias.split /C/
    tempitemalias = tempitemalias[0].to_s
    tempitemalias2 = tempitemalias.chop
    puts tempitemalias2
    @collegiatecolor = @print.split /Spirit /
    @collegiatecolor = @collegiatecolor[1].to_s


    while uuidlastdigit < 3

      if uuidlastdigit == 0


        open('thirtyone', 'ftc2wtppsdb01')
        query("SELECT top 1 inv_code FROM [thirtyone].[dbo].[table_product] where inv_code like '#{tempitemalias2}%'and IsKitHeader = '1' and description like '%#{@collegiatecolor}%'")
        close
        @data.flatten

        @kitheader = @data[0]
        puts @kitheader

        @OrderSystemUUID = "#{@ordnum.rstrip}" + "|1|" +"#{uuidlastdigit}"
        @Barcode = "#{@ordnum.rstrip}" + "#{barcodelastnum}"


        wrksheet.Cells(@collegerows, "A").value = @OrderSystemUUID
        wrksheet.Cells(@collegerows, "B").value = 1
        wrksheet.Cells(@collegerows, "C").value = @kitheader
        wrksheet.Cells(@collegerows, "D").value = 1
        wrksheet.Cells(@collegerows, "F").value = 1
        wrksheet.Cells(@collegerows, "E").value = lineseq
        wrksheet.Cells(@collegerows, "G").value = @Barcode
        wrksheet.Cells(@collegerows, "M").value = 'true'
        wrksheet.Cells(@collegerows, "N").value = 'false'
        wrksheet.Cells(@collegerows, "O").value = @kitheader
        wrksheet.Cells(@collegerows, "AD").value = "Entered"


        uuidlastdigit = uuidlastdigit + 1
        barcodelastnum = barcodelastnum + 1
        lineseq = lineseq + 1
        @collegerows = @collegerows + 1


      end

      if uuidlastdigit == 1

        @OrderSystemUUID = "#{@ordnum.rstrip}" + "|1|" +"#{uuidlastdigit}"
        @Barcode = "#{@ordnum.rstrip}" + "#{barcodelastnum}"


        wrksheet.Cells(@collegerows, "A").value = @OrderSystemUUID
        wrksheet.Cells(@collegerows, "B").value = 5
        wrksheet.Cells(@collegerows, "C").value = @itemalias
        wrksheet.Cells(@collegerows, "D").value = 1
        wrksheet.Cells(@collegerows, "E").value = lineseq
        wrksheet.Cells(@collegerows, "F").value = @quantity
        wrksheet.Cells(@collegerows, "G").value = @Barcode
        wrksheet.Cells(@collegerows, "H").value = @option4

        open('31StagingStore', 'cmh2wdsql02')
        query("select style_out from [31StagingStore].[dbo].[FontXlat]where sku = ('#{tempitemalias}') and style_in = '#{@option2.rstrip}'")
        close
        @data.flatten

        @embroiderfontstyle = @data[0].to_s

        if @option3 =~ /all caps/

          @option3 = @option3.split("(")
          @option3 = @option3[0].rstrip


        end


        wrksheet.Cells(@collegerows, "I").value = @embroiderfontstyle
        wrksheet.Cells(@collegerows, "J").value = @option3
        wrksheet.Cells(@collegerows, "M").value = 'false'
        wrksheet.Cells(@collegerows, "N").value = 'true'
        wrksheet.Cells(@collegerows, "O").value = @kitheader
        wrksheet.Cells(@collegerows, "AD").value = "Entered"
        wrksheet.Cells(@collegerows, "AE").value = @school
        wrksheet.Cells(@collegerows, "AF").value = "White"

        #Query below is using current barcode to grab ItemID that maps to parentITEMID when UUID == 2


        #open('31StagingStore', 'cmh2wdsql02')
        #query("SELECT ItemIDFROM [31StagingStore].[dbo].[Item]where ItemAlias like '%C%' and Barcode = '#{@Barcode}'")
        #close
        #@data.flatten
        #
        #@ParentItemID = @data[0]

        uuidlastdigit = uuidlastdigit + 1
        barcodelastnum = barcodelastnum + 1
        lineseq = lineseq + 1
        @collegerows = @collegerows + 1



      end

      if uuidlastdigit == 2


        @OrderSystemUUID = "#{@ordnum.rstrip}" + "|1|" +"#{uuidlastdigit}"
        @Barcode = "#{@ordnum.rstrip}" + "#{barcodelastnum}"

        open('thirtyone', 'ftc2wtppsdb01')
        query("select inv_code FROM [thirtyone].[dbo].[table_product] where inv_code like '4146%' and description like '%#{@school}%' and inv_code not like '%B%'")
        close
        @data.flatten


        @itemalias = @data[0]


        wrksheet.Cells(@collegerows, "A").value = @OrderSystemUUID
        wrksheet.Cells(@collegerows, "B").value = 1
        wrksheet.Cells(@collegerows, "C").value = @itemalias
        wrksheet.Cells(@collegerows, "D").value = 1
        wrksheet.Cells(@collegerows, "E").value = lineseq
        wrksheet.Cells(@collegerows, "F").value = @quantity
        wrksheet.Cells(@collegerows, "G").value = @Barcode
        wrksheet.Cells(@collegerows, "M").value = 'false'
        wrksheet.Cells(@collegerows, "N").value = 'true'
        wrksheet.Cells(@collegerows, "O").value = @kitheader
        wrksheet.Cells(@collegerows, "AD").value = "Entered"

        #Need Solution for ParentItemID since it is not available until Orders Enter Staging Store
        #wrksheet.Cells(@collegerows, "AG").value = @ParentItemID


        uuidlastdigit = uuidlastdigit + 1
        barcodelastnum = barcodelastnum + 1
        lineseq = lineseq + 1
        @collegerows = @collegerows + 1


      end

      #@collegerows = @collegerows + 1


    end


  end


  if @personalizationtypeid == "14"

    #switch from using style_in because Icon it no text does not have style_in. Using personalization type to pull back style_out

    open('31StagingStore', 'cmh2wdsql02')
    query("select style_out from [31StagingStore].[dbo].[FontXlat]where sku = ('#{@itemalias}') and PersonalizationTypeID = '#{@personalizationtypeid}'")
    close
    @data.flatten


    @embroiderfontstyle = @data[0].to_s


    wrksheet.Cells(@rows, "I").value = @embroiderfontstyle
    wrksheet.Cells(@rows, "M").value = "false"
    wrksheet.Cells(@rows, "N").value = "false"
    wrksheet.Cells(@rows, "AD").value = "Entered"


    @designstyle = @embroiderfontstyle
    @designstyle = @designstyle.split (/- */)


    @designstyle = @designstyle[1].to_s


    wrksheet.Cells(@rows, "AE").value = "#{@option1}"+"-"+"#{@designstyle}"
    wrksheet.Cells(@rows, "AF").value = @option2

  end


  wrkbook.save
  excel.quit


end