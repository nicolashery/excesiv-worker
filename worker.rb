require 'bundler/setup'
require 'uri'
require 'jmongo'

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
    task_oid = doc['_id']
    puts "Processing task #{task_oid}"
    task_id = BSON::ObjectId.from_string(task_oid)
    sleep 5 # Simulate some processing time
    @results.insert({'task' => {'_id' => task_id},
                    'message' => doc['message'].reverse})
    puts "Done with task #{task_oid}"
  end

  def run
    puts "Excesiv worker starting..."
    # First loop is to go to the end of the collection
    puts "Going to end of collection"
    cursor = Mongo::Cursor.new(@tasks, :timeout => false, :tailable => true)
    # Note: we can't test on cursor.next here, because there might be some
    # residual 'poison docs' in the collection that return nil, so it would
    # stop us from going to the end of the collection
    cursor.count.times do
      cursor.next
    end
    # Second loop keeps going, waiting for new data or Ctrl+C
    puts "Listening for new data"
    loop do
      # Note: when we are at the end of the collection, jmongo will insert
      # a 'poison doc' and cursor.next will return nil
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