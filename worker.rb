require 'bundler/setup'
require 'uri'
require 'mongo'

STDOUT.sync = true # Write in real-time

# Envrionment variables for config
mongodb_uri = ENV['MONGOLAB_URI'] || 'mongodb://localhost/excesiv'

class Worker

  def initialize(mongodb_uri)
    # Parse mongdb_uri
    uri = URI.parse(mongodb_uri)
    # Connect to database
    connection = Mongo::Connection.from_uri(mongodb_uri)
    db = connection.db(uri.path.gsub(/^\//, ''))
    # Collections
    @tasks = db.collection('tasks')
    @results = db.collection('results')
    # File system
    @fs = Mongo::Grid.new(db)
    @fs_meta = db.collection('fs.files')
  end

  def process_task(doc)
    task_id = doc['_id']
    template = doc['template']
    attachment_filename = doc['attachment_filename']
    puts "Processing task #{task_id}, template #{template}"
    sleep 3 # Simulate some processing time
    # Get the correct template from MongoDB
    template_meta = @fs_meta.find_one({'filename' => template, 
                            'label' => 'template'})
    template_id = template_meta['_id']
    f = @fs.get(template_id)
    # Save result file back to database
    file_id = @fs.put(f, :filename => "result_#{template}",
                      :content_type  =>  f.content_type,
                      :label => 'result',
                      :attachment_filename => attachment_filename)
    # Notify result queue that we are finished, with link to file
    @results.insert({'task' => {'_id' => task_id},
                    'file' => {'_id' => file_id},
                    'message' => doc['message'].reverse})
    puts "Done with task #{task_id}"
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
