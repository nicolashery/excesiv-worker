require 'bundler/setup'
require 'uri'
require 'mongo'

class Excesiv::Worker

  # Config
  APP_DEBUG = false
  CAPPED_COLLECTION_SIZE = 8000000
  CAPPED_COLLECTION_MAX = 100

  def initialize
    # Core object to use Excel
    @xs = Excesiv::Core.new
    # Parse mongdb_uri
    mongodb_uri = ENV['MONGOLAB_URI'] || ENV['MONGODB_URL'] || 
              'mongodb://localhost/excesiv'
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
      wb = @xs.open_wb(f_in)
      @xs.write_wb(wb, data)
      # Save result file back to database
      f_out = @fs_store.open("result_#{template}", 'w', 
                            :content_type  =>  content_type,
                            :label => 'result',
                            :attachment_filename => attachment_filename)
      file_id = f_out.files_id
      @xs.save_wb(wb, f_out)
      # Result is a link to generated file
      result = {'task_id' => task_id, 'file_id' => file_id}
    elsif task_type == 'read'
      file_id = doc['file_id']
      f_in = @fs.get(file_id)
      wb = @xs.open_wb(f_in)
      data = @xs.read_wb(wb)
      result = {'task_id' => task_id, 'data' => data}
    end
    # Send result back in MongoDB queue
    @results.insert(result)
  end

  def run
    STDOUT.sync = true # Write in real-time
    puts "Excesiv v#{Excesiv::VERSION}"
    puts "Worker starting..."
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
          task_id = doc['_id']
          puts "Processing task #{task_id}"
          # In production, keep worker alive if processing task fails
          if APP_DEBUG
            process_task doc
            puts "Done with task #{task_id}"
          else
            begin
              process_task doc
              puts "Done with task #{task_id}"
            rescue => exception
              # Send error message back to result queue
              result = {'task_id' => task_id, 
                        'error' => exception.message}
              @results.insert(result)
              puts "Error with task #{task_id}"
              puts exception.message
              puts exception.backtrace
            end
          end
        end
      else
        sleep 1
      end
    end
  end

end # class Worker

