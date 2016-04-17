# GlobalExcelReader [![Build Status](https://travis-ci.org/woahdae/simple_xlsx_reader.svg?branch=master)](https://travis-ci.org/woahdae/simple_xlsx_reader)

This is a global reader for s3 excel files, as well as both xml, xlsx, and xls files. This grabs the data, the styles, and other important attributes, and outputs an xml file you can format in your views.

The intial structure is based off of the gem 'simple_xlsx_reader'.

This is *not* a rewrite of excel in Ruby. Font styles, for
example, are parsed to determine whether a cell is a number or a date,
then forgotten. We just want to get the data, and get out!

## Usage

### Summary:


### Load Errors


### More

## Installation

Add this line to your application's Gemfile:

    gem 'global_excel_reader'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install global_excel_reader

## Versioning

This project follows [semantic versioning 1.0](http://semver.org/spec/v1.0.0.html)

## Contributing

Remember to write tests, think about edge cases, and run the existing
suite.

Follow this path for feature requests:

1. Fork this project
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
