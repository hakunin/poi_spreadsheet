
Poi Spreadheet
=========

[![Gem Version](https://badge.fury.io/rb/poi_spreadsheet.svg)](http://badge.fury.io/rb/poi_spreadsheet)

This is a simple spreadsheet wrapper around the powerfull Apache POI for Ruby.

At this point, simple reading and writing is supported. If you need more, adding functionality is very simple.
Just look up the Java documentation for Apache POI and add wrapper for what you need (make sure to send a pull request too).


Installation
---------------------
In your Gemfile:
```ruby
gem 'poi_spreadsheet'
```

Usage
---------------------
```ruby
# load your file
book = PoiSpreadsheet.load 'my_sheet.xls'

# read the first sheet
10.times { |x| 
  10.times { |y| 
    puts "#{x}: #{y} #{book.sheets.first[x][y].inspect}" 
  }
}
```

Write some data
```ruby
book.sheets.first[0][0] = 1
book.sheets.first[0][1] = "Hello World"
```

Save the document

```ruby
book.save # save into original
book.save 'file_name.xls'
```

Manipulate sheets

```ruby
book.sheets.length # => 1
book.clone_sheet(0)
book.sheets.length # => 2
```

```ruby
book.create_sheet 'name'
```

```ruby
book.remove_sheet_at 0
```


IDEAS
-----
 - support cell names (A3..)
 - manipulate cell style
 - merge cells
 - read, write and execute formulas
 - add support for xlsx


Problems? Issues? Want to help out?
---------------------
Just submit a Gihub issue or pull request! 

The MIT License (MIT)
---------------------
Copyright (c) 2014 Michal Hantl

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
