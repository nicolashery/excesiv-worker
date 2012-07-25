require 'bundler/setup'

require_relative 'poi'

class Excesiv

  # Open workbook from file object or file path
  def open_wb(f_in)
    # If file path given, open file
    if f_in.kind_of? String
      f_in = File.open(f_in, 'r')
    end
    # Convert Ruby IO object to Java InputStream
    # http://jruby.org/apidocs/org/jruby/util/IOInputStream.html
    f_in = org.jruby.util.IOInputStream.new(f_in)
    # Generate workbook from template
    wb = Poi::XSSFWorkbook.new(f_in)
    wb
  end

  # Save workbook to file object or file path
  def save_wb(wb, f_out)
    # If file path given, open file
    if f_out.kind_of? String
      f_out = File.open(f_out, 'w')
    end
    # Convert Ruby IO object to Java OutputStream
    # http://jruby.org/apidocs/org/jruby/util/IOOutputStream.html
    f_out = org.jruby.util.IOOutputStream.new(f_out)
    wb.write(f_out)
  end

  # Return the number of the first row defined in the workbook names
  def get_first_row(wb)
    # Rows start at 0 in API, at 1 in Excel, so substract 1
    first_row = Integer(wb.getName('first_row').getRefersToFormula()) - 1
    first_row
  end

  # Return a dictionary with the instructions defined by the named ranges in 
  # the workbook. Use defined first row of first worksheet for styles and 
  # formulas.
  def get_instructions(wb)
    instructions = {
      'w' => {
        'header' => [],
        'data' => [],
        'formula' => []
      },
      'r' => {
        'header' => [],
        'data' => [],
        'formula' => []
      }
    }
    ws = wb.getSheetAt(0)
    first_row = get_first_row(wb)
    for i in 0..(wb.getNumberOfNames() - 1)
      name_obj = wb.getNameAt(i)
      comment = name_obj.getComment()
      if comment
        tags = comment.split
      else
        tags = []
      end
      # Check that name object is an instruction
      if not (tags & ['w', 'r']).empty? and \
          not (tags & ['header', 'data', 'formula']).empty?
        instruction = {}
        instruction['name'] = name_obj.getNameName()
        # Capture target column numbers 
        arearef = Poi::AreaReference.new(name_obj.getRefersToFormula())
        cellrefs = arearef.getAllReferencedCells()
        columns = cellrefs.map{|c| c.getCol()}
        instruction['columns'] = columns
        # For header cells only
        if tags.include?('header')
          # Capture target row number
          rows = cellrefs.map{|c| c.getRow()}
          instruction['row'] = rows[0]
        # For row cells only
        elsif not (tags & ['data', 'formula']).empty?
          # Style: use first cell of instruction column range
          j = columns[0]
          cell = ws.getRow(first_row).getCell(j)
          instruction['style'] = cell.getCellStyle()
          # For formula row cells only
          if tags.include?('formula')
            # Formula is assumed to be the same for all columns of instruction,
            # use first cell
            if cell.getCellType() != 3 # Blank cell type
              instruction['formula'] = cell.getCellFormula()
            else
              instruction['formula'] = ''
            end
          end
          # For instruction with more than 1 column, allow last cell to be 
          # styled differently
          if columns.length > 1
            j = columns[-1]
            cell = ws.getRow(first_row).getCell(j)
            instruction['style_last'] = cell.getCellStyle()
          end
        end
        # Add instruction to correct places in instructions map
        if tags.include?('w')
          if tags.include?('header')
            instructions['w']['header'] << instruction
          elsif tags.include?('data')
            instructions['w']['data'] << instruction
          elsif tags.include?('formula')
            instructions['w']['formula'] << instruction
          end
        end
        if tags.include?('r')
          if tags.include?('header')
            instructions['r']['header'] << instruction
          elsif tags.include?('data')
            instructions['r']['data'] << instruction
          elsif tags.include?('formula')
            instructions['r']['formula'] << instruction
          end
        end
      end
    end
    instructions
  end

  def test
    f_in = File.open('test.xlsx', 'r')
    wb = open_wb(f_in)
    puts get_instructions(wb)
  end

end

#xs = Excesiv.new

#xs.test

