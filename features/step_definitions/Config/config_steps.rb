require 'win32ole'
require 'watir'

Given(/^I Place an Order$/) do


  @rows = 80
  count = 1
  @collegerows = 2

  while count < 9


    browser
    login
    retailord
    randproduct
    randpersonalizationoption
    randoption1
    randoption2
    randoption3
    randoption4
    randoption5
    textbox1
    textbox2
    submitord
    datatoexcelDBOrder
    datatoexcelDBItem
    datatoexcelDBpersonalizations
    datatoexcelDBpulseorders


    count = count + 1
    @rows = @rows + 1

  end


end


Given(/^I Extract Excel Data$/) do


  @rows = 2


  while  @rows <5

    extractdataexcelDBOrder
    comparedatadborder
    extractdataexceldbitem
    comparedatadbitem
    extractdataexceldbitemcollegiate
    comparedatadbitemcollegiate


    @rows = @rows + 1

  end

end

Given(/^I Extract Excel Pulse Data$/) do

  @rows = 2

  while @rows <5

    extractdataDBPers
    compareDBPers
    extractDBpersorders
    compareDBpersOrders


    @rows = @rows + 1

  end





end


def open(database, source)

  # Open ADO connection to the SQL Server database
  connection_string = "Provider=SQLOLEDB.1;"
  connection_string << "Persist Security Info=False;"
  connection_string << "User ID=31GIFTS\awarns;"
  connection_string << "Trusted_Connection=yes;"
  connection_string << "Initial Catalog=#{database};"
  connection_string << "Data Source=#{source};"

  #puts connection_string

  @connection = WIN32OLE.new('ADODB.Connection')
  @connection.Open(connection_string)


end

def query(sql)


  # Create an instance of an ADO Recordset
  recordset = WIN32OLE.new('ADODB.Recordset')
  # Open the recordset, using an SQL statement and the
  # existing ADO connection
  recordset.Open(sql, @connection)
  # Create and populate an array of field names
  @fields = []

  recordset.Fields.each do |field|
    @fields << field.Name
  end
  begin
    # Move to the first record/row, if any exist
    recordset.MoveFirst
    # Grab all records
    @data = recordset.GetRows
  rescue
    @data = []
  end
  recordset.Close
  # An ADO Recordset's GetRows method returns an array
  # of columns, so we'll use the transpose method to
  # convert it to an array of rows
  @data = @data.flatten

end

def close

  @connection.Close

end


def quitexcel



  sleep(2)

end
