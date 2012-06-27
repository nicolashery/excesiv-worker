require 'bundler/setup'
require 'uri'
require 'mongo'

require_relative 'poi'

STDOUT.sync = true # Write in real-time

# Envrionment variables for config
mongodb_uri = ENV['MONGOLAB_URI'] || 'mongodb://localhost/excesiv'

class Worker

  def initialize(mongodb_uri)
    # Parse mongdb_uri
    uri = URI.parse(mongodb_uri)
    # Connect to database
    connection = Mongo::Connection.from_uri(mongodb_uri)
    @db = connection.db(uri.path.gsub(/^\//, ''))
    # Collections
    @tasks = @db.collection('tasks')
    @results = @db.collection('results')
    # File system
    @fs = Mongo::Grid.new(@db)
    @fs_meta = @db.collection('fs.files')
    @fs_store = Mongo::GridFileSystem.new(@db) # Used for new file
  end

  def process_task(doc)
    task_id = doc['_id']
    template = doc['template']
    attachment_filename = doc['attachment_filename']
    data = doc['data']
    puts "Processing task #{task_id}, template #{template}"
    #sleep 3 # Simulate some processing time
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
    generate_workbook(wb, data)
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
    # Notify result queue that we are finished, with link to file
    @results.insert({'task' => {'_id' => task_id},
                    'file' => {'_id' => file_id}})
    puts "Done with task #{task_id}"
  end

  # Populate workbook with data
  def generate_workbook(wb, data={})
    message = data['message'] or 'Hello World!'

    ws = wb.getSheetAt(0)

    row = ws.getRow(0)
    cell = row.getCell(1)
    cell.setCellValue(message)

    row = ws.getRow(2)
    cell = row.getCell(1)
    cell.setCellValue(message.reverse)
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
