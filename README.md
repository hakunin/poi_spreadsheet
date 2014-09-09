
=======
Poi Spreadheet
=========

This is a simple spreadsheet wrapper around the powerfull Apache POI for Ruby.

At this point, simple reading and writing is supported. If you need more, adding functionality is very simple.
Just look up the Java documentation for Apache POI and add wrapper for what you need (make sure to send a pull request too).


Installation
---------------------
In your Gemfile:
```ruby
gem 'poi_spreadsheet'
```
Getting Started
---------------------
```ruby
# load your file
book = PoiSpreadsheet.load 'xls-templates/timesheet.xls'; 

# read the first sheet
10.times { |x| 
  10.times { |y| 
    puts "#{x}: #{y} #{book.sheets.first.last[x][y].inspect}" 
  }
}
```

IDEAS
-----

 - add support for adding/removing/modifying sheets
 - add support for xlsx
 - add support for styling


Problems? Issues? Want to help out?
---------------------
Just submit a Gihub issue or pull request! 

The MIT License (MIT)
---------------------
Copyright (c) 2014 Pathgather

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
