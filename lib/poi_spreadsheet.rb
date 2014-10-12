require 'rjb'

class PoiSpreadsheet
  

  def self.init
    apache_poi_path = File.dirname(__FILE__)+'/../apache/poi-3.10.1-20140818.jar'
    Rjb::load(apache_poi_path, ['-Xmx512M'])

    @cell_class = cell_class = Rjb::import('org.apache.poi.hssf.usermodel.HSSFCell')


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



  class Workbook
  
    attr_accessor :j_book

    def self.load file
      @file_name = file

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

    # Get sheets
    def sheets
      @sheets ||= begin
        sheets = []
        self.j_book.getNumberOfSheets.times { |i|
          j_sheet = j_book.getSheetAt(i)
          sheet = Worksheet.from_sheet(j_sheet)
          sheet.book = self
          name = j_book.getSheetName(i)
          sheets << sheet
        }
        sheets
      end
    end

    def create_sheet name
      self.sheets << Worksheet.from_sheet(j_book.createSheet(name))
    end

    def clone_sheet index
      self.sheets << Worksheet.from_sheet(j_book.cloneSheet(index))
    end

    def remove_sheet_at index
      j_book.removeSheetAt(index)
      @sheets.delete_at(index)
    end

    # Get sheet by name
    def [](k)
      sheets[k]
    end

    def save file_name = @file_name
      @file_output_class ||= Rjb::import('java.io.FileOutputStream')
      out = @file_output_class.new(file_name);

      begin
        j_book.write(out)
      ensure
        out.close();
      end
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

    def name
      j_sheet.getSheetName
    end

    def name= name
      j_book = j_sheet.getWorkbook
      j_book.setSheetName(j_book.getSheetIndex(j_sheet), name)
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

      def []= col, value
        cell = j_row.getCell(col)
        cell.setCellValue(value)
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



