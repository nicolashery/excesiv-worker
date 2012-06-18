require 'bundler/setup'
require 'uri'
require 'mongo'

STDOUT.sync = true # Write in real-time

# Envrionment variables for config
mongodb_uri = ENV['MONGODB_URI'] || 'mongodb://localhost/excesiv'

class Worker

  def initialize(mongodb_uri='mongodb://localhost/excesiv')
    # Parse mongdb_uri
    uri = URI.parse(mongodb_uri)
    # Connect to database
    connection = Mongo::Connection.from_uri(mongodb_uri)
    db = connection.db(uri.path.gsub(/^\//, ''))
    # Collections
    @tasks = db.collection('tasks')
    @results = db.collection('results')
  end

  def process_task(doc)
    task_id = doc['_id']
    puts "Processing task #{task_id}"
    sleep 3 # Simulate some processing time
    @results.insert({'task' => {'_id' => task_id},
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