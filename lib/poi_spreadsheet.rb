require 'rjb'

class PoiSpreadsheet
  

  def self.init
    apache_poi_path = File.dirname(__FILE__)+'/../apache/poi-3.10.1-20140818.jar'
    Rjb::load(apache_poi_path, ['-Xmx512M'])

    @cell_class = cell_class = Rjb::import('org.apache.poi.hssf.usermodel.HSSFCell')

    # Java classes import 
    @file_class = Rjb::import('java.io.FileOutputStream')

    Rjb::import('org.apache.poi.hssf.usermodel.HSSFCreationHelper')
    Rjb::import('org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator')

    @cell_reference_class = Rjb::import('org.apache.poi.hssf.util.CellReference')
    # You can import all java classes that you need
    @loaded = true
  end

  def self.cell_class; @cell_class; end


  def self.load file
    puts 'Hello!'
    unless @loaded
      init
    end
    Workbook.load file
  end


  # for xls use org.apache.poi.hssf.extractor.ExcelExtractor

=begin 

  Example java code

  Workbook wb = new HSSFWorkbook(fis); //or new XSSFWorkbook("/somepath/test.xls")
  Sheet sheet = wb.getSheetAt(0);
  FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

  // suppose your formula is in B3
  CellReference cellReference = new CellReference("B3"); 
  Row row = sheet.getRow(cellReference.getRow());
  Cell cell = row.getCell(cellReference.getCol()); 

  if (cell!=null) {
    switch (evaluator.evaluateFormulaCell(cell)) {
      case Cell.CELL_TYPE_BOOLEAN:
          System.out.println(cell.getBooleanCellValue());
          break;
      case Cell.CELL_TYPE_NUMERIC:
          System.out.println(cell.getNumericCellValue());
          break;
      case Cell.CELL_TYPE_STRING:
          System.out.println(cell.getStringCellValue());
          break;
      case Cell.CELL_TYPE_BLANK:
          break;
      case Cell.CELL_TYPE_ERROR:
          System.out.println(cell.getErrorCellValue());
          break;

      // CELL_TYPE_FORMULA will never occur
      case Cell.CELL_TYPE_FORMULA: 
          break;
    }
  }
=end


  class Workbook
  
    attr_accessor :j_book

    def self.load file
      @workbook_class = Rjb::import('org.apache.poi.hssf.usermodel.HSSFWorkbook')
      @poifs_class = Rjb::import('org.apache.poi.poifs.filesystem.POIFSFileSystem')
      @file_input_class = Rjb::import('java.io.FileInputStream')
      @file_input = @file_input_class.new(file)

      book = new

      fs = @poifs_class.new(@file_input)
      book.j_book = @workbook_class.new(fs)
      book
    end

    def initialize
      @sheets = nil
    end

    # Get sheet names
    def sheets
      @sheets ||= begin
        sheets = {}
        self.j_book.getNumberOfSheets.times { |i|
          j_sheet = j_book.getSheetAt(i)
          sheet = Worksheet.from_sheet(j_sheet)
          sheet.book = self
          name = j_book.getSheetName(i)
          sheets[name] = sheet
        }
        sheets
      end
    end

    # Get sheet by name
    def [](k)
      sheets[k]
    end

    def _evaluator
      @_evaluator ||= j_book.getCreationHelper.createFormulaEvaluator
    end
  end

  class Worksheet

    attr_accessor :j_sheet
    attr_accessor :book


    def initialize
      @rows = {}
    end

    # get cell
    def [](row)
      @rows[row] ||= begin
        j_row = j_sheet.getRow(row)
        row = Row.from_row j_row
        row.sheet = self
        row
      end
    end

    # set cell
    def set(x, y)
      
    end

    def self.from_sheet j_sheet
      sheet = new
      sheet.j_sheet = j_sheet
      sheet
    end

    class Row

      attr_accessor :j_row
      attr_accessor :sheet

      def self.symbol_type(constant)
        @types ||= begin
          cell = ::PoiSpreadsheet.cell_class
          {
            cell.CELL_TYPE_BOOLEAN => :boolean,
            cell.CELL_TYPE_NUMERIC => :numeric,
            cell.CELL_TYPE_STRING => :string,
            cell.CELL_TYPE_BLANK => :blank,
            cell.CELL_TYPE_ERROR => :error,
            cell.CELL_TYPE_FORMULA => :formula,
          }
        end
        @types[constant]
      end
      
      def [] col
        unless cell = j_row.getCell(col)
          return nil
        end

        #type = self.class.symbol_type(sheet.book._evaluator.evaluateFormulaCell(cell))
        type = self.class.symbol_type(cell.getCellType())
        
        case type
        when :boolean
          cell.getBooleanCellValue()
        when :numeric
          cell.getNumericCellValue()
        when :string
          cell.getStringCellValue()
        when :blank
          nil
        when :error
          cell.getErrorCellValue()
        when :formula
          cell.getNumericCellValue()
        end
      end

      def self.from_row j_row
        row = new
        row.j_row = j_row
        row
      end

    end

  end

end



