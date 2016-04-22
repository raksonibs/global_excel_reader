# GlobalExcelReader [![Build Status](https://travis-ci.org/raksonibs/global_excel_reader.svg?branch=master)](https://travis-ci.org/raksonibs/global_excel_reader)

This is a global reader for s3 excel files, as well as both xml, xlsx, and xls files. This grabs the data and outputs a hash you can format in your views.

The intial structure is based off of the gem 'simple_xlsx_reader'.

All Styles are forgotten, and only data is parsed.

## Usage
  All you have to do is pass a file to this gem and it will do the rest. Extensions it supports are xls, csv, xml, and xls.

  The gem throws nil for other files. No need to interrupt your flow.

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

Follow this path for feature requests:

1. Fork this project
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
