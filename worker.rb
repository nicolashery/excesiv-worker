require 'bundler/setup'
require 'uri'
require 'mongo'

require_relative 'poi'

STDOUT.sync = true # Write in real-time

# Envrionment variables for config
mongodb_uri = ENV['MONGOLAB_URI'] || ENV['MONGOLAB_URI'] || 
              'mongodb://localhost/excesiv'

class Worker

  # Config
  CAPPED_COLLECTION_SIZE = 1000000
  CAPPED_COLLECTION_MAX = 3

  def initialize(mongodb_uri)
    # Parse mongdb_uri
    uri = URI.parse(mongodb_uri)
    # Connect to database
    connection = Mongo::Connection.from_uri(mongodb_uri)
    @db = connection.db(uri.path.gsub(/^\//, ''))
    # Create capped collections unless they exists already
    if @db.collection_names.include? 'tasks'
      @tasks = @db.collection('tasks')
    else
      @tasks = @db.create_collection('tasks', :capped => true, 
                                :autoIndexId => true,
                                :size => CAPPED_COLLECTION_SIZE, 
                                :max => CAPPED_COLLECTION_MAX)
      # Insert a dummy document just in case because some drivers have trouble 
      # with empty capped collections
      @tasks.insert({'init' => true})
    end
    if @db.collection_names.include? 'results'
      @results = @db.collection('results')
    else
      @results = @db.create_collection('results', :capped => true, 
                                :autoIndexId => true,
                                :size => CAPPED_COLLECTION_SIZE, 
                                :max => CAPPED_COLLECTION_MAX)
      @results.insert({'init' => true})
    end
    # File system
    @fs = Mongo::Grid.new(@db)
    @fs_meta = @db.collection('fs.files')
    @fs_store = Mongo::GridFileSystem.new(@db) # Used for new file
  end

  def process_task(doc)
    task_id = doc['_id']
    task_type = doc['type']
    puts "Processing task #{task_id}"
    if task_type == 'write'
      template = doc['template']
      attachment_filename = doc['attachment_filename']
      data = doc['data']
      # Get the correct template from the database
      template_meta = @fs_meta.find_one({'filename' => template, 
                              'label' => 'template'})
      template_id = template_meta['_id']
      f_in = @fs.get(template_id)
      content_type= f_in.content_type
      # Convert Ruby IO object to Java InputStream
      # http://jruby.org/apidocs/org/jruby/util/IOInputStream.html
      f_in = org.jruby.util.IOInputStream.new(f_in)
      # Generate workbook from template
      wb = Poi::XSSFWorkbook.new(f_in)
      f_in.close()
      write_workbook(wb, data)
      # Save result file back to database
      f_out = @fs_store.open("result_#{template}", 'w', 
                            :content_type  =>  content_type,
                            :label => 'result',
                            :attachment_filename => attachment_filename)
      file_id = f_out.files_id
      # Convert Ruby IO object to Java OutputStream
      # http://jruby.org/apidocs/org/jruby/util/IOOutputStream.html
      f_out = org.jruby.util.IOOutputStream.new(f_out)
      wb.write(f_out)
      f_out.close()
      # Result is a link to generated file
      result = {'task_id' => task_id, 'file_id' => file_id}
    elsif task_type == 'read'
      file_id = doc['file_id']
      f_in = @fs.get(file_id)
      f_in = org.jruby.util.IOInputStream.new(f_in)
      wb = Poi::XSSFWorkbook.new(f_in)
      f_in.close()
      data = read_workbook(wb)
      result = {'task_id' => task_id, 'data' => data}
    end
    # Send result back in MongoDB queue
    @results.insert(result)
    puts "Done with task #{task_id}"
  end

  # Populate workbook with data
  def write_workbook(wb, data={})
    message = data['message'] or 'Hello World!'

    ws = wb.getSheetAt(0)

    row = ws.getRow(0)
    cell = row.getCell(1)
    cell.setCellValue(message)

    row = ws.getRow(2)
    cell = row.getCell(1)
    cell.setCellValue(message.reverse)
  end

  # Read data from workbook
  def read_workbook(wb)
    data = {}
    ws = wb.getSheetAt(0)
    row = ws.getRow(0)
    cell = row.getCell(1)
    # Optional response key in data is sent back to web client
    data['response'] = "#{cell}".reverse
    data
  end

  def run
    puts "Excesiv worker starting..."
    # First loop is to go to the end of the collection
    puts "Going to end of collection"
    cursor = Mongo::Cursor.new(@tasks, :timeout => false, :tailable => true)
    cursor.count.times do
      cursor.next
    end
    # Second loop keeps going, waiting for new data or Ctrl+C
    puts "Listening for new data"
    loop do
      if doc = cursor.next
        # Check that task hasn't been picked up by another worker yet
        # and if not, set the 'assigned' indicator to true
        # (find_and_modify returns nil if not found, the doc otherwise)
        is_not_assigned = @tasks.find_and_modify( \
          {'query' => {'_id' => doc['_id'], 'assigned' => false}, 
           'update' => {'$set' => {'assigned' => true}}, 
           'new' => true})
        if is_not_assigned
          process_task doc
        end
      else
        sleep 1
      end
    end
  end

end # class Worker

worker = Worker.new(mongodb_uri)

worker.run
