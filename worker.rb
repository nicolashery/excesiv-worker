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
    sleep 5 # Simulate some processing time
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
        process_task doc
      else
        sleep 1
      end
    end
  end

end # class Worker

worker = Worker.new(mongodb_uri)

worker.run