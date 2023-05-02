require 'nokogiri'
require 'open-uri'
require 'uri'
require 'caxlsx'
require 'win32ole'

# Ask the user to enter a search term
puts "Enter a search term:"
search_term = gets.chomp

# Ask the user to choose between current and sold listings
puts "Do you want to search for (c)urrent or (s)old listings?"
listing_type = gets.chomp.downcase

# Construct the eBay search results URL using the search term and listing type
url = "https://www.ebay.com/sch/i.html?_nkw=#{URI::DEFAULT_PARSER.escape(search_term)}"
if listing_type == 's'
  url += '&LH_Sold=1'
end

# Use Nokogiri to parse the HTML from the URL
html = URI.open(url).read
doc = Nokogiri::HTML(html)

# Find all the listings on the page
listings = doc.css('.s-item')

# Initialize the data array
data = [['Item Name', 'Condition', 'Sold Price', 'Shipping']]

# Loop through each listing and extract the data we want
listings.each do |listing|
  # Extract the title, price, and shipping cost
  title = listing.css('.s-item__title').text.strip
  condition = listing.css('.SECONDARY_INFO').text.strip
  price = listing.css('.s-item__price').text.gsub('US $', '$').gsub(' to ', '-')
  shipping = listing.css('.s-item__shipping.s-item__logisticsCost').text.strip

  # Split the shipping text into type and amount
  if shipping == 'Free shipping'
    shipping_type, shipping_amount = 'Free Shipping', nil
  elsif shipping.include?('to')
    shipping_type, shipping_amount = 'Range', shipping.gsub(' shipping','').strip
  else
    shipping_type, shipping_amount = 'Amount Charged', shipping.gsub('+','').strip
  end

  # Add the data to the array
  data << [title, condition, price, "#{shipping_type} - #{shipping_amount}"]
end

# Create the Excel file
Axlsx::Package.new do |p|
  p.workbook.add_worksheet(:name => "eBay Listings") do |sheet|
    data.each do |row|
      sheet.add_row row
    end
  end

  # Ask the user if they want to save the file
  puts "Do you want to save the data to a file? (y/n)"
  save_file = gets.chomp.downcase

  if save_file == 'y'
    # Open Windows Explorer to let the user choose the save location and file name
    shell = WIN32OLE.new('Shell.Application')
    save_dialog = shell.BrowseForFolder(0, "Select the folder to save the file in", 0, 0)
    if save_dialog != nil
      save_folder = save_dialog.Items().Item().Path
      puts "Enter a file name:"
      file_name = gets.chomp
      file_path = File.join(save_folder, file_name + '.xlsx') # include the .xlsx extension

      # Save the Excel file
      p.serialize(file_path)

      puts "Data saved to #{file_path}"
    end
  else
    puts "Data not saved"
  end
end
