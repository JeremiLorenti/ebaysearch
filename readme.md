# eBay Search Using Ruby

This program allows you to search eBay for either current or sold listings using a search term you specify. It then extracts data from the search results, such as item name, condition, sold price, and shipping information, and saves it to an Excel file.

## Installation

To run this program, you'll need to have Ruby and the following gems installed:

- `nokogiri`
- `open-uri`
- `uri`
- `caxlsx`
- `win32ole`

You can install these gems by running the following command in your terminal:

```sh
gem install nokogiri open-uri uri caxlsx win32ole
```

## Usage

To use this program, simply run the `ebaysearch.rb` file from the terminal:

```sh
ruby ebaysearch.rb
```

You'll be prompted to enter a search term, choose between current or sold listings, and whether you want to save the data to an Excel file. If you choose to save the data, you'll be prompted to choose a folder and enter a file name.

## License

This program is released under the [MIT License](https://opensource.org/licenses/MIT).

## Credits

This program was created by Jeremi Lorenti.