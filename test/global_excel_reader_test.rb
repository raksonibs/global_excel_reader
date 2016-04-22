require_relative 'test_helper'
require 'time'

GER = GlobalExcelReader

describe GlobalExcelReader do
  let(:sesame_street_blog_file) { File.join(File.dirname(__FILE__),
                                            'sesame_street_blog.xlsx') }
  let(:animal_network_csv_file) { File.join(File.dirname(__FILE__),
                                            'animal_network.csv') }
  let(:animal_network_xls_file) { File.join(File.dirname(__FILE__),
                                            'animal_network.xls') }
  let(:animal_network_xml_file) { File.join(File.dirname(__FILE__),
                                            'animal_network_2.xml') }

  let(:subject) { GlobalExcelReader::Document.new(sesame_street_blog_file) }
  let(:subject_csv) { GlobalExcelReader::Document.new(animal_network_csv_file) }
  let(:subject_xls) { GlobalExcelReader::Document.new(animal_network_xls_file) }
  let(:subject_xml) { GlobalExcelReader::Document.new(animal_network_xml_file) }

  describe '#to_hash' do
    it 'reads an xlsx file into a hash of {[sheet name] => [data]}' do
      subject.to_hash.must_equal({
        "Authors"=>
          [["Name", "Occupation"],
           ["Big Bird", "Teacher"]],

        "Posts"=>
          [["Author Name", "Title", "Body", "Created At", "Comment Count", "URL"],
           ["Big Bird", "The Number 1", "The Greatest", Time.parse("2002-01-01 11:00:00 UTC"), 1, GER::Hyperlink.new("http://www.example.com/hyperlink-function", "This uses the HYPERLINK() function")],
           ["Big Bird", "The Number 2", "Second Best", Time.parse("2002-01-02 14:00:00 UTC"), 2, GER::Hyperlink.new("http://www.example.com/hyperlink-gui", "This uses the hyperlink GUI option")],
           ["Big Bird", "Formula Dates", "Tricky tricky", Time.parse("2002-01-03 14:00:00 UTC"), 0, nil],
           ["Empty Eagress", nil, "The title, date, and comment have types, but no values", nil, nil, nil]]
      })
    end
  end

  # describe '#convert_csv' do
  #   it 'reads a csv file into a hash of {[sheet name] => [data]}' do
  #     subject_csv.to_hash.must_equal({"Project Claim Form"=>[["Animal Network Project Claim Form", nil], [nil], [nil], ["Animal Network  Project Information - Combined", nil, nil, nil, nil, nil, nil, nil, nil], ["Name of Animal Network:", "oskar test", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Name:", "1", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Type", nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Contract start date:", "2015-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Contract completion date:", "2016-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Event dates:", "2012-12-13, 2013,13,13", nil, nil, nil, nil, nil, nil, nil, nil], ["Other relevant Information:", "whatever yo", nil, nil, nil, nil, nil, nil, nil, nil], [], ["Item", "YoPlease attach copies of the actual receipts for the 100% expenses listed.\r\rPlease only fill out the actual expenses at the conclusion of the project.\r\rAllowable Expenses\rProject costs may include contractors, consultants, incremental office expenses, licenses and fees, marketing and recruitment, venue and audio visual event costs, transportation, research, education, materials and supplies. \r\rProject costs including capital costs, overhead or alcohol will not be eligible for reimbursement.", "Never", "Justein", "2015 01 03", "Today", nil, nil, nil, nil, nil, nil, nil, nil], ["Cats", "100", "200", "0", "75", "-125", nil, nil, nil, nil, nil, nil, nil, nil], ["Dogs", "200", "100", "0", "225", "125", nil, nil, nil, nil, nil, nil, nil, nil], ["Birds", "300", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["Maps", "0", "400", "0", "17.25", "-382.75", nil, nil, nil, nil, nil, nil, nil, nil], ["Mereetings", "0", "0", "0", "33", "33", nil, nil, nil, nil, nil, nil, nil, nil], ["Toho", "150", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["More cats costs, as related to delivery of events and marketing activities.", "0", "0", "0", "75", "75", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Totals (Internal Use)", "750", "700", "0", "425.25", "-274.75", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Anticipated Cash Flow Requests", "Note: If there are any changes, please make the anticipated project cash flow request changes for the whole year, not to exceed 75% of the total expense before tax.", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["From April 2016 - March 2017", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["REPORTING PERIOD / COST CATEGORY Please fill in the amounts that you will be requesting from WCAN reimbursement for in relation to your projects.on a monthly basis.", "Apr", "May", "June", "July", "Aug", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow Across All Projects (Most Updated)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow (Per Budget - Internal Use Only)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, "0", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Woot", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0.75", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Yo", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil], ["Notes: Create a formula where the monthly cash flow is an aggregate amount from all of the projects for each Animal Network.", nil, nil, nil, nil, nil, nil, nil, nil], ["Sept", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "TOTAL forecasted"], ["Total Requested", nil, "Total Budgeted", "Total Spent", "Variance", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, "0", "0", "0", "0", "0", "0", "0", nil], ["0", nil, nil, "0", nil, nil, "0", nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", nil, "0", "0"], []], "Sheet2"=>[["Animal Network Project Claim Form", nil], [nil], [nil], ["Animal Network  Project Information - Combined", nil, nil, nil, nil, nil, nil, nil, nil], ["Name of Animal Network:", "oskar test", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Name:", "2", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Type", nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Contract start date:", "2015-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Contract completion date:", "2016-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Event dates:", "2012-12-13, 2013,13,13", nil, nil, nil, nil, nil, nil, nil, nil], ["Other relevant Information:", "whatever yo", nil, nil, nil, nil, nil, nil, nil, nil], [], ["Item", "DogsPlease attach copies of the actual receipts for the 100% expenses listed.\r\rPlease only fill out the actual expenses at the conclusion of the project.\r\rAllowable Expenses\rProject costs may include contractors, consultants, incremental office expenses, licenses and fees, marketing and recruitment, venue and audio visual event costs, transportation, research, education, materials and supplies. \r\rProject costs including capital costs, overhead or alcohol will not be eligible for reimbursement.", "cats", "LOl", "Whatsup", "Variance", nil, nil, nil, nil, nil, nil, nil, nil], ["Cats", "1", "10", "0", "75", "65", nil, nil, nil, nil, nil, nil, nil, nil], ["Dogs", "2", "11", "0", "225", "214", nil, nil, nil, nil, nil, nil, nil, nil], ["Birds", "3", "12", "0", "0", "-12", nil, nil, nil, nil, nil, nil, nil, nil], ["Maps", "4", "13", "0", "17.25", "4.25", nil, nil, nil, nil, nil, nil, nil, nil], ["Mereetings", "5", "14", "0", "33", "19", nil, nil, nil, nil, nil, nil, nil, nil], ["Toho", "6", "15", "0", "0", "-15", nil, nil, nil, nil, nil, nil, nil, nil], ["More cats costs, as related to delivery of events and marketing activities.", "7", "16", "0", "75", "59", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Totals (Internal Use)", "28", "91", "0", "425.25", "334.25", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Anticipated Cash Flow Requests", "Note: If there are any changes, please make the anticipated project cash flow request changes for the whole year, not to exceed 75% of the total expense before tax.", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["From April 2016 - March 2017", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["REPORTING PERIOD / COST CATEGORY Please fill in the amounts that you will be requesting from WCAN reimbursement for in relation to your projects.on a monthly basis.", "Apr", "May", "June", "July", "Aug", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow Across All Projects (Most Updated)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow (Per Budget - Internal Use Only)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, "0", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Woot", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0.75", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Yo", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil], ["Notes: Create a formula where the monthly cash flow is an aggregate amount from all of the projects for each Animal Network.", nil, nil, nil, nil, nil, nil, nil, nil], ["Sept", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "TOTAL forecasted"], ["Total Requested", nil, "Total Budgeted", "Total Spent", "Variance", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, "0", "0", "0", "0", "0", "0", "0", nil], ["0", nil, nil, "0", nil, nil, "0", nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", nil, "0", "0"], []]})
  #   end
  # end

  describe '#convert_xml' do
    it 'reads an xml file into a hash of {[sheet name] => [data]}' do
      subject_xml.to_hash.must_equal({"Project Claim Form"=>[["Animal Network Project Claim Form", nil], [nil], [nil], ["Animal Network  Project Information - Combined", nil, nil, nil, nil, nil, nil, nil, nil], ["Name of Animal Network:", "oskar test", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Name:", "1", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Type", nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Contract start date:", "2015-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Contract completion date:", "2016-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Event dates:", "2012-12-13, 2013,13,13", nil, nil, nil, nil, nil, nil, nil, nil], ["Other relevant Information:", "whatever yo", nil, nil, nil, nil, nil, nil, nil, nil], [], ["Item", "YoPlease attach copies of the actual receipts for the 100% expenses listed.\r\rPlease only fill out the actual expenses at the conclusion of the project.\r\rAllowable Expenses\rProject costs may include contractors, consultants, incremental office expenses, licenses and fees, marketing and recruitment, venue and audio visual event costs, transportation, research, education, materials and supplies. \r\rProject costs including capital costs, overhead or alcohol will not be eligible for reimbursement.", "Never", "Justein", "2015 01 03", "Today", nil, nil, nil, nil, nil, nil, nil, nil], ["Cats", "100", "200", "0", "75", "-125", nil, nil, nil, nil, nil, nil, nil, nil], ["Dogs", "200", "100", "0", "225", "125", nil, nil, nil, nil, nil, nil, nil, nil], ["Birds", "300", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["Maps", "0", "400", "0", "17.25", "-382.75", nil, nil, nil, nil, nil, nil, nil, nil], ["Mereetings", "0", "0", "0", "33", "33", nil, nil, nil, nil, nil, nil, nil, nil], ["Toho", "150", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["More cats costs, as related to delivery of events and marketing activities.", "0", "0", "0", "75", "75", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Totals (Internal Use)", "750", "700", "0", "425.25", "-274.75", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Anticipated Cash Flow Requests", "Note: If there are any changes, please make the anticipated project cash flow request changes for the whole year, not to exceed 75% of the total expense before tax.", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["From April 2016 - March 2017", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["REPORTING PERIOD / COST CATEGORY Please fill in the amounts that you will be requesting from WCAN reimbursement for in relation to your projects.on a monthly basis.", "Apr", "May", "June", "July", "Aug", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow Across All Projects (Most Updated)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow (Per Budget - Internal Use Only)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, "0", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Woot", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0.75", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Yo", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil], ["Notes: Create a formula where the monthly cash flow is an aggregate amount from all of the projects for each Animal Network.", nil, nil, nil, nil, nil, nil, nil, nil], ["Sept", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "TOTAL forecasted"], ["Total Requested", nil, "Total Budgeted", "Total Spent", "Variance", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, "0", "0", "0", "0", "0", "0", "0", nil], ["0", nil, nil, "0", nil, nil, "0", nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", nil, "0", "0"], []], "Sheet2"=>[["Animal Network Project Claim Form", nil], [nil], [nil], ["Animal Network  Project Information - Combined", nil, nil, nil, nil, nil, nil, nil, nil], ["Name of Animal Network:", "oskar test", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Name:", "2", nil, nil, nil, nil, nil, nil, nil, nil], ["Project Type", nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Contract start date:", "2015-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Contract completion date:", "2016-06-01", nil, nil, nil, nil, nil, nil, nil, nil], ["Event dates:", "2012-12-13, 2013,13,13", nil, nil, nil, nil, nil, nil, nil, nil], ["Other relevant Information:", "whatever yo", nil, nil, nil, nil, nil, nil, nil, nil], [], ["Item", "DogsPlease attach copies of the actual receipts for the 100% expenses listed.\r\rPlease only fill out the actual expenses at the conclusion of the project.\r\rAllowable Expenses\rProject costs may include contractors, consultants, incremental office expenses, licenses and fees, marketing and recruitment, venue and audio visual event costs, transportation, research, education, materials and supplies. \r\rProject costs including capital costs, overhead or alcohol will not be eligible for reimbursement.", "cats", "LOl", "Whatsup", "Variance", nil, nil, nil, nil, nil, nil, nil, nil], ["Cats", "1", "10", "0", "75", "65", nil, nil, nil, nil, nil, nil, nil, nil], ["Dogs", "2", "11", "0", "225", "214", nil, nil, nil, nil, nil, nil, nil, nil], ["Birds", "3", "12", "0", "0", "-12", nil, nil, nil, nil, nil, nil, nil, nil], ["Maps", "4", "13", "0", "17.25", "4.25", nil, nil, nil, nil, nil, nil, nil, nil], ["Mereetings", "5", "14", "0", "33", "19", nil, nil, nil, nil, nil, nil, nil, nil], ["Toho", "6", "15", "0", "0", "-15", nil, nil, nil, nil, nil, nil, nil, nil], ["More cats costs, as related to delivery of events and marketing activities.", "7", "16", "0", "75", "59", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Totals (Internal Use)", "28", "91", "0", "425.25", "334.25", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Anticipated Cash Flow Requests", "Note: If there are any changes, please make the anticipated project cash flow request changes for the whole year, not to exceed 75% of the total expense before tax.", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["From April 2016 - March 2017", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["REPORTING PERIOD / COST CATEGORY Please fill in the amounts that you will be requesting from WCAN reimbursement for in relation to your projects.on a monthly basis.", "Apr", "May", "June", "July", "Aug", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow Across All Projects (Most Updated)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], ["Projected Cash Flow (Per Budget - Internal Use Only)", "0", "0", "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, "0", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Woot", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0.75", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil], ["Yo", "0", nil, "0", "0", "0", nil, nil, nil, nil, nil, nil, nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil], ["Notes: Create a formula where the monthly cash flow is an aggregate amount from all of the projects for each Animal Network.", nil, nil, nil, nil, nil, nil, nil, nil], ["Sept", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "TOTAL forecasted"], ["Total Requested", nil, "Total Budgeted", "Total Spent", "Variance", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, "0", "0", "0", "0", "0", "0", "0", nil], ["0", nil, nil, "0", nil, nil, "0", nil, nil], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", "0", "0", "0"], [nil, nil, nil, nil, nil, nil, nil, nil, nil], ["0", "0", "0", "0", "0", "0", nil, "0", "0"], []]})
    end
  end

  describe '#convert_xls' do
    it 'reads an xls file into a hash of {[sheet name] => [data]}' do
      subject_xls.to_hash.must_equal({"Project Claim Form"=>
  [[" Animal Network Project Claim Form",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Animal Network  Project Information - Combined",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Name of Animal Network:",
    "oskar test",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Project Name:", 2.0, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Project Type", "", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Contract start date:",
    "2015-06-01",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Contract completion date:",
    "2016-06-01",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Event dates: ",
    "2012-12-13, 2013,13,13",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Other relevant Information:",
    "whatever yo",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Item",
    nil,
    "Dogs",
    "cats",
    "LOl ",
    "Whatsup",
    "Variance ",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Cats", nil, 1.0, 10.0, 0.0, 75.0, 65.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Dogs", nil, 2.0, 11.0, 0.0, 225.0, 214.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Birds", nil, 3.0, 12.0, 0.0, 0.0, -12.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Maps", nil, 4.0, 13.0, 0.0, 17.25, 4.25, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Mereetings", nil, 5.0, 14.0, 0.0, 33.0, 19.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Toho", nil, 6.0, 15.0, 0.0, 0.0, -15.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["More cats costs, as related to delivery of events and marketing activities.",
    nil,
    7.0,
    16.0,
    0.0,
    75.0,
    59.0,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Totals (Internal Use)",
    nil,
    28.0,
    91.0,
    0.0,
    425.25,
    334.25,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Anticipated Cash Flow Requests",
    "Note: If there are any changes, please make the anticipated project cash flow request changes for the whole year, not to exceed 75% of the total expense before tax.",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["From April 2016 - March 2017",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["REPORTING PERIOD / COST CATEGORY ",
    nil,
    "Apr",
    "May",
    "June",
    "July",
    "Aug",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Projected Cash Flow Across All Projects (Most Updated)",
    nil,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Projected Cash Flow (Per Budget - Internal Use Only)",
    nil,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, 0.0, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Woot", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [0.75, nil, 0.0, nil, 0.0, 0.0, 0.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Yo", nil, 0.0, nil, 0.0, 0.0, 0.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Notes: Create a formula where the monthly cash flow is an aggregate amount from all of the projects for each Animal Network.",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    "Sept",
    "Oct",
    "Nov",
    "Dec",
    "Jan",
    "Feb",
    "Mar",
    "TOTAL forecasted ",
    nil],
   [nil,
    nil,
    "Total Requested",
    nil,
    "Total Budgeted",
    "Total Spent",
    "Variance",
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, nil, nil, 0.0, nil, nil, 0.0, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, nil, 0.0, 0.0]],
 "animal_network.csv"=>
  [[" Animal Network Project Claim Form",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Animal Network  Project Information - Combined",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Name of Animal Network:",
    "oskar test",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Project Name:", 2.0, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Project Type", "", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Contract start date:",
    "2015-06-01",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Contract completion date:",
    "2016-06-01",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Event dates: ",
    "2012-12-13, 2013,13,13",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Other relevant Information:",
    "whatever yo",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Item",
    nil,
    "Dogs",
    "cats",
    "LOl ",
    "Whatsup",
    "Variance ",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Cats", nil, 1.0, 10.0, 0.0, 75.0, 65.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Dogs", nil, 2.0, 11.0, 0.0, 225.0, 214.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Birds", nil, 3.0, 12.0, 0.0, 0.0, -12.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Maps", nil, 4.0, 13.0, 0.0, 17.25, 4.25, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Mereetings", nil, 5.0, 14.0, 0.0, 33.0, 19.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Toho", nil, 6.0, 15.0, 0.0, 0.0, -15.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["More cats costs, as related to delivery of events and marketing activities.",
    nil,
    7.0,
    16.0,
    0.0,
    75.0,
    59.0,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Totals (Internal Use)",
    nil,
    28.0,
    91.0,
    0.0,
    425.25,
    334.25,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Anticipated Cash Flow Requests",
    "Note: If there are any changes, please make the anticipated project cash flow request changes for the whole year, not to exceed 75% of the total expense before tax.",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["From April 2016 - March 2017",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["REPORTING PERIOD / COST CATEGORY ",
    nil,
    "Apr",
    "May",
    "June",
    "July",
    "Aug",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Projected Cash Flow Across All Projects (Most Updated)",
    nil,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   ["Projected Cash Flow (Per Budget - Internal Use Only)",
    nil,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil, nil, nil, nil, 0.0, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Woot", nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [0.75, nil, 0.0, nil, 0.0, 0.0, 0.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Yo", nil, 0.0, nil, 0.0, 0.0, 0.0, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   ["Notes: Create a formula where the monthly cash flow is an aggregate amount from all of the projects for each Animal Network.",
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil],
   [nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    nil,
    "Sept",
    "Oct",
    "Nov",
    "Dec",
    "Jan",
    "Feb",
    "Mar",
    "TOTAL forecasted ",
    nil],
   [nil,
    nil,
    "Total Requested",
    nil,
    "Total Budgeted",
    "Total Spent",
    "Variance",
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    0.0,
    nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, nil, nil, 0.0, nil, nil, 0.0, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
   [nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil, nil],
   [nil, nil, nil, nil, nil, nil, nil, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, nil, 0.0, 0.0]]})
    end
  end

  describe GlobalExcelReader::Document::Mapper do
    let(:described_class) { GlobalExcelReader::Document::Mapper }

    describe '::cast' do
      it 'reads type s as a shared string' do
        described_class.cast('1', 's', nil, :shared_strings => ['a', 'b', 'c']).
          must_equal 'b'
      end

      it 'reads type inlineStr as a string' do
        described_class.cast('the value', nil, 'inlineStr').
          must_equal 'the value'
      end

      it 'reads date styles' do
        described_class.cast('41505', nil, :date).
          must_equal Date.parse('2013-08-19')
      end

      it 'reads time styles' do
        described_class.cast('41505.77083', nil, :time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads date_time styles' do
        described_class.cast('41505.77083', nil, :date_time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads number types styled as dates' do
        described_class.cast('41505', 'n', :date).
          must_equal Date.parse('2013-08-19')
      end

      it 'reads number types styled as times' do
        described_class.cast('41505.77083', 'n', :time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads less-than-zero complex number types styled as times' do
        described_class.cast('6.25E-2', 'n', :time).
          must_equal Time.parse('1899-12-30 01:30:00 UTC')
      end

      it 'reads number types styled as date_times' do
        described_class.cast('41505.77083', 'n', :date_time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'raises when date-styled values are not numerical' do
        lambda { described_class.cast('14 is not a valid date', nil, :date) }.
          must_raise(ArgumentError)
      end

      describe "with the url option" do
        let(:url) { "http://www.example.com/hyperlink" }
        it 'creates a hyperlink with a string type' do
          described_class.cast("A link", 'str', :string, url: url).
            must_equal GER::Hyperlink.new(url, "A link")
        end

        it 'creates a hyperlink with a shared string type' do
          described_class.cast("2", 's', nil, shared_strings: ['a','b','c'], url: url).
            must_equal GER::Hyperlink.new(url, 'c')
        end
      end
    end

    describe '#shared_strings' do
      let(:xml) do
        GlobalExcelReader::Document::Xml.new.tap do |xml|
          xml.shared_strings = Nokogiri::XML(File.read(
            File.join(File.dirname(__FILE__), 'shared_strings.xml') )).remove_namespaces!
        end
      end

      subject { described_class.new(xml) }

      it 'parses strings formatted at the cell level' do
        subject.shared_strings[0..2].must_equal ['Cell A1', 'Cell B1', 'My Cell']
      end

      it 'parses strings formatted at the character level' do
        subject.shared_strings[3..5].must_equal ['Cell A2', 'Cell B2', 'Cell Fmt']
      end
    end

    describe '#style_types' do
      let(:xml) do
        GlobalExcelReader::Document::Xml.new.tap do |xml|
          xml.styles = Nokogiri::XML(File.read(
            File.join(File.dirname(__FILE__), 'styles.xml') )).remove_namespaces!
        end
      end

      let(:mapper) do
        GlobalExcelReader::Document::Mapper.new(xml)
      end

      it 'reads custom formatted styles (numFmtId >= 164)' do
        mapper.style_types[1].must_equal :date_time
        mapper.custom_style_types[164].must_equal :date_time
      end

      # something I've seen in the wild; don't think it's correct, but let's be flexible.
      it 'reads custom formatted styles given an id < 164, but not explicitly defined in the SpreadsheetML spec' do
        mapper.style_types[2].must_equal :date_time
        mapper.custom_style_types[59].must_equal :date_time
      end
    end

    describe '#last_cell_label' do

      let(:generic_style) do
          Nokogiri::XML(
            <<-XML
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <cellXfs count="1">
                <xf numFmtId="0" />
              </cellXfs>
            </styleSheet>
            XML
          ).remove_namespaces!
      end

      # Note, this is not a valid sheet, since the last cell is actually D1 but
      # the dimension specifies C1. This is just for testing.
      let(:sheet) do
        Nokogiri::XML(
          <<-XML
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <dimension ref="A1:C1" />
            <sheetData>
              <row>
                <c r='A1' s='0'>
                  <v>Cell A</v>
                </c>
                <c r='C1' s='0'>
                  <v>Cell C</v>
                </c>
                <c r='D1' s='0'>
                  <v>Cell D</v>
                </c>
              </row>
            </sheetData>
          </worksheet>
          XML
        ).remove_namespaces!
      end

      let(:empty_sheet) do
        Nokogiri::XML(
          <<-XML
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <dimension ref="A1" />
            <sheetData>
            </sheetData>
          </worksheet>
          XML
        ).remove_namespaces!
      end

      let(:xml) do
        GlobalExcelReader::Document::Xml.new.tap do |xml|
          xml.sheets = [sheet]
          xml.styles = generic_style
        end
      end

      subject { described_class.new(xml) }

      it 'uses /worksheet/dimension if available' do
        subject.last_cell_label(sheet).must_equal 'C1'
      end

      it 'uses the last header cell if /worksheet/dimension is missing' do
        sheet.xpath('/worksheet/dimension').remove
        subject.last_cell_label(sheet).must_equal 'D1'
      end

      it 'returns "A1" if the dimension is just one cell' do
        subject.last_cell_label(empty_sheet).must_equal 'A1'
      end

      it 'returns "A1" if the sheet is just one cell, but /worksheet/dimension is missing' do
        sheet.at_xpath('/worksheet/dimension').remove
        subject.last_cell_label(empty_sheet).must_equal 'A1'
      end
    end

    describe '#column_letter_to_number' do
      let(:subject) { described_class.new }

      [ ['A',   1    ],
        ['B',   2    ],
        ['Z',   26   ],
        ['AA',  27   ],
        ['AB',  28   ],
        ['AZ',  52   ],
        ['BA',  53   ],
        ['BZ',  78   ],
        ['ZZ',  702  ],
        ['AAA', 703  ],
        ['AAZ', 728  ],
        ['ABA', 729  ],
        ['ABZ', 754  ],
        ['AZZ', 1378 ],
        ['ZZZ', 18278] ].each do |(letter, number)|
        it "converts #{letter} to #{number}" do
          subject.column_letter_to_number(letter).must_equal number
        end
      end
    end

    describe "parse errors" do
      after do
        GlobalExcelReader.configuration.catch_cell_load_errors = false
      end

      let(:xml) do
        GlobalExcelReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
            <<-XML
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <dimension ref="A1:A1" />
              <sheetData>
                <row>
                  <c r='A1' s='0'>
                    <v>14 is a date style; this is not a date</v>
                  </c>
                </row>
              </sheetData>
            </worksheet>
            XML
          ).remove_namespaces!]

          # s='0' above refers to the value of numFmtId at cellXfs index 0
          xml.styles = Nokogiri::XML(
            <<-XML
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <cellXfs count="1">
                <xf numFmtId="14" />
              </cellXfs>
            </styleSheet>
            XML
          ).remove_namespaces!
        end
      end

      it 'raises if configuration.catch_cell_load_errors' do
        GlobalExcelReader.configuration.catch_cell_load_errors = false

        lambda { described_class.new(xml).parse_sheet('test', xml.sheets.first, nil) }.
          must_raise(GlobalExcelReader::CellLoadError)
      end

      it 'records a load error if not configuration.catch_cell_load_errors' do
        GlobalExcelReader.configuration.catch_cell_load_errors = true

        sheet = described_class.new(xml).parse_sheet('test', xml.sheets.first, nil)
        sheet.load_errors[[0,0]].must_include 'invalid value for Float'
      end
    end

    describe "missing numFmtId attributes" do

      let(:xml) do
        GlobalExcelReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
                            <<-XML
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <dimension ref="A1:A1" />
              <sheetData>
                <row>
                  <c r='A1' s='s'>
                    <v>some content</v>
                  </c>
                </row>
              </sheetData>
            </worksheet>
                        XML
                        ).remove_namespaces!]

          xml.styles = Nokogiri::XML(
              <<-XML
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

            </styleSheet>
          XML
          ).remove_namespaces!
        end
      end

      before do
        @row = described_class.new(xml).parse_sheet('test', xml.sheets.first, nil).rows[0]
      end

      it 'continues even when cells are missing numFmtId attributes ' do
        @row[0].must_equal 'some content'
      end

    end

    describe 'parsing types' do
      let(:xml) do
        GlobalExcelReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
            <<-XML
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dimension ref="A1:G1" />
                <sheetData>
                  <row>
                    <c r='A1' s='0'>
                      <v>Cell A1</v>
                    </c>

                    <c r='C1' s='1'>
                      <v>2.4</v>
                    </c>
                    <c r='D1' s='1' />

                    <c r='E1' s='2'>
                      <v>30687</v>
                    </c>
                    <c r='F1' s='2' />

                    <c r='G1' t='inlineStr' s='0'>
                      <is><t>Cell G1</t></is>
                    </c>

                    <c r='H1' s='0'>
                      <f>HYPERLINK("http://www.example.com/hyperlink-function", "HYPERLINK function")</f>
                      <v>HYPERLINK function</v>
                    </c>

                    <c r='I1' s='0'>
                      <v>GUI-made hyperlink</v>
                    </c>
                  </row>
                </sheetData>

                <hyperlinks>
                  <hyperlink ref="I1" id="rId1"/>
                </hyperlinks>
              </worksheet>
            XML
          ).remove_namespaces!]

          # s='0' above refers to the value of numFmtId at cellXfs index 0,
          # which is in this case 'General' type
          xml.styles = Nokogiri::XML(
            <<-XML
              <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <cellXfs count="1">
                  <xf numFmtId="0" />
                  <xf numFmtId="2" />
                  <xf numFmtId="14" />
                </cellXfs>
              </styleSheet>
            XML
          ).remove_namespaces!

          # Although not a "type" or "style" according to xlsx spec,
          # it sure could/should be, so let's test it with the rest of our
          # typecasting code.
          xml.sheet_rels = [Nokogiri::XML(
            <<-XML
              <Relationships>
                <Relationship
                  Id="rId1"
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                  Target="http://www.example.com/hyperlink-gui"
                  TargetMode="External"
                />
              </Relationships>
            XML
          ).remove_namespaces!]

        end
      end

      before do
        @row = described_class.new(xml).parse_sheet('test', xml.sheets.first, xml.sheet_rels.first).rows[0]
      end

      it "reads 'Generic' cells as strings" do
        @row[0].must_equal "Cell A1"
      end

      it "reads empty 'Generic' cells as nil" do
        @row[1].must_equal nil
      end

      # We could expand on these type tests, but really just a couple
      # demonstrate that it's wired together. Type-specific tests should go
      # on #cast

      it "reads floats" do
        @row[2].must_equal 2.4
      end

      it "reads empty floats as nil" do
        @row[3].must_equal nil
      end

      it "reads dates" do
        @row[4].must_equal Date.parse('Jan 6, 1984')
      end

      it "reads empty date cells as nil" do
        @row[5].must_equal nil
      end

      it "reads strings formatted as inlineStr" do
        @row[6].must_equal 'Cell G1'
      end

      it "reads hyperlinks created via HYPERLINK()" do
        @row[7].must_equal(
          GER::Hyperlink.new(
            "http://www.example.com/hyperlink-function", "HYPERLINK function"))
      end

      it "reads hyperlinks created via the GUI" do
        @row[8].must_equal(
          GER::Hyperlink.new(
            "http://www.example.com/hyperlink-gui", "GUI-made hyperlink"))
      end
    end

    describe 'parsing documents with blank rows' do
      let(:xml) do
        GlobalExcelReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
            <<-XML
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dimension ref="A1:D7" />
                <sheetData>
                <row r="2" spans="1:1">
                  <c r="A2" s="0">
                    <v>0</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="B4" s="0">
                    <v>1</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="C5" s="0">
                    <v>2</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="D7" s="0">
                    <v>3</v>
                  </c>
                </row>
                </sheetData>
              </worksheet>
            XML
          ).remove_namespaces!]

          xml.styles = Nokogiri::XML(
            <<-XML
              <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <cellXfs count="1">
                  <xf numFmtId="0" />
                </cellXfs>
              </styleSheet>
            XML
          ).remove_namespaces!
        end
      end

      before do
        @rows = described_class.new(xml).parse_sheet('test', xml.sheets.first, nil).rows
      end

      it "reads row data despite gaps in row numbering" do
        @rows.must_equal [
          [nil,nil,nil,nil],
          ["0",nil,nil,nil],
          [nil,nil,nil,nil],
          [nil,"1",nil,nil],
          [nil,nil,"2",nil],
          [nil,nil,nil,nil],
          [nil,nil,nil,"3"]
        ]
      end
    end

  end
end
