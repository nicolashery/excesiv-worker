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

  # Helper to set cell value, turning nil values into empty strings
  # POI seems to have trouble with setting cells to nil values, 
  # and sometimes crashes with the following error:
  # DateUtil.java:94:in `getExcelDate': java.lang.NullPointerException
  def set_cell_value(cell, value)
    if value.nil?
      value = ''
    end
    cell.setCellValue(value)
  end

  # Helper to return cell value based on cell type
  # For formula cells, the last calcluated result by Excel is returned
  # Date values will be returned as a Float, and will have to be converted
  # using Excel's method for storing dates (see Excel documentation)
  def get_cell_value(cell)
    if not cell
      value = nil
    else
      cell_type = cell.getCellType()
      # Formula cells
      if cell_type == Poi::CELL_TYPE_FORMULA
        # Cached formula result will be a String,
        # need to convert it based on type
        cell_type = cell.getCachedFormulaResultType()
        value = cell.getRawValue()
        case cell_type
        when Poi::CELL_TYPE_NUMERIC
          # Apparently a formula returning 0 gives a nil value, fix it
          if value.nil?
            value = 0
          end
          value = Float(value)
        when Poi::CELL_TYPE_STRING
          # Do nothing
        when Poi::CELL_TYPE_BLANK
          value = ''
        when Poi::CELL_TYPE_BOOLEAN
          # Formula result will be '1' or '0'
          value = value == '1' ? true : false
        when Poi::CELL_TYPE_ERROR
          # Do nothing
        else
          value = nil
        end
      # Value cells
      else
        case cell_type
        when Poi::CELL_TYPE_NUMERIC
          value = cell.getNumericCellValue()
        when Poi::CELL_TYPE_STRING
          value = cell.getStringCellValue()
        when Poi::CELL_TYPE_BLANK
          value = ''
        when Poi::CELL_TYPE_BOOLEAN
          value = cell.getBooleanCellValue()
        when Poi::CELL_TYPE_ERROR
          value = cell.getErrorCellString()
        else
          value = nil
        end
      end
    end
    value
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
        'rows' => []
      },
      'r' => {
        'header' => [],
        'rows' => []
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
          instruction['type'] = 'header'
          # Capture target row number
          rows = cellrefs.map{|c| c.getRow()}
          instruction['row'] = rows[0]
        # For row cells only
        elsif not (tags & ['data', 'formula']).empty?
          # Style: use first cell of instruction column range
          j = columns[0]
          cell = ws.getRow(first_row).getCell(j)
          instruction['style'] = cell.getCellStyle()
          if tags.include?('formula')
            instruction['type'] = 'formula'
            # Formula is assumed to be the same for all columns of instruction,
            # use first cell
            if cell.getCellType() == 2 # Formula cell type
              instruction['formula'] = cell.getCellFormula()
            else
              instruction['formula'] = ''
            end
          else
            instruction['type'] = 'data'
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
          elsif not (tags & ['data', 'formula']).empty?
            instructions['w']['rows'] << instruction
          end
        end
        if tags.include?('r')
          if tags.include?('header')
            instructions['r']['header'] << instruction
          elsif not (tags & ['data', 'formula']).empty?
            instructions['r']['rows'] << instruction
          end
        end
      end
    end
    instructions
  end

  # Fill template workbook with data, use first worksheet
  def write_wb(wb, data={'header'=>{}, 'rows'=>[]})
    ws = wb.getSheetAt(0)
    # Get instructions from named ranges in workbook
    instructions = get_instructions(wb)
    # Header
    for instruction in instructions['w']['header']
      name = instruction['name']
      if data['header'].has_key?(name)
        values = data['header'][name]
        # Convert single values to list to be able to loop
        values = [*values]
        row = ws.getRow(instruction['row'])
        columns = instruction['columns']
        # Fill cells with values
        columns.zip(values).each do |j, value|
          cell = row.getCell(j)
          set_cell_value(cell, value)
        end
      end
    end
    # Rows
    i = get_first_row(wb)
    for row_data in data['rows']
      row = ws.getRow(i) ? ws.getRow(i) : ws.createRow(i)
      for instruction in instructions['w']['rows']
        columns = instruction['columns']
        style = instruction['style']
        # Fill cells with values or formula
        if instruction['type'] == 'data'
          values = row_data[instruction['name']] 
          values = [*values]
          columns.zip(values).each do |j, value|
            cell = row.getCell(j) ? row.getCell(j) : row.createCell(j)
            set_cell_value(cell, value)
            cell.setCellStyle(style)
          end
        else
          formula = instruction['formula']
          columns.each do |j|
            cell = row.getCell(j) ? row.getCell(j) : row.createCell(j)
            cell.setCellFormula(formula)
            cell.setCellStyle(style)
          end
        end
        # For instructions on multiple columns, 
        # last cell can have different style 
        if columns.length > 1
          j = columns[-1]
          style = instruction['style_last']
          cell = row.getCell(j)
          cell.setCellStyle(style)
        end
      end
      # Next row
      i = i + 1
    end
    # Force recalulation of all formulas
    wb.setForceFormulaRecalculation(true)
    wb
  end
  
  # Read data from workbook, use first worksheet
  def read_wb(wb)
    data = {'header' => {}, 'rows' => []}
    ws = wb.getSheetAt(0)
    # Get instructions from named ranges in workbook
    instructions = get_instructions(wb)
    # Header
    for instruction in instructions['r']['header']
      name = instruction['name']
      row = ws.getRow(instruction['row'])
      columns = instruction['columns']
      # Read cell values
      values = []
      for j in columns
        cell = row.getCell(j)
        values << get_cell_value(cell)
      end
      # If there is only one columns, convert back to single value
      if values.length == 1
        values = values[0]
      end
      data['header'][name] = values
    end
    # Rows
    first_row = get_first_row(wb)
    last_row = ws.getLastRowNum()
    for i in (first_row..last_row)
      row = ws.getRow(i)
      row_data = {}
      for instruction in instructions['r']['rows']
        name = instruction['name']
        columns = instruction['columns']
        # Read cells values
        # Note: the "formula r" instruction is equivalent to "data r"
        values = []
        for j in columns
          cell = row.getCell(j)
          values << get_cell_value(cell)
        end
        if values.length == 1
          values = values[0]
        end
        row_data[name] = values
      end
      data['rows'] << row_data
    end
    data
  end

  def test_cellvalue
    STDOUT.sync = true
    f_in = File.open('test_cellvalue.xlsx', 'r')
    f_in = org.jruby.util.IOInputStream.new(f_in)
    wb = Poi::XSSFWorkbook.new(f_in)
    ws = wb.getSheetAt(0)
    puts "Test value cells"
    for i in (1..6)
      cell = ws.getRow(i).getCell(2)
      val = get_cell_value(cell)
      puts "cell #{i}"
      cell_type = cell ? cell.getCellType() : nil
      puts "cell type #{cell_type}"
      puts "value #{val}"
      puts "value class #{val.class}"
    end
    puts "Test formula cells"
    for i in (1..6)
      cell = ws.getRow(i).getCell(3)
      val = get_cell_value(cell)
      puts "cell #{i}"
      cell_type = cell ? cell.getCellType() : nil
      puts "cell type #{cell_type}"
      puts "cached formula result type #{cell.getCachedFormulaResultType()}"
      puts "value #{val}"
      puts "value class #{val.class}"
    end
  end
end

if __FILE__ == $0
  xs = Excesiv.new
  xs.test
end

