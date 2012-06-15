
def run
  puts "Excesiv worker starting..."
  i = 1
  while(true)
    begin
      puts "Listening for new data #{i}"
      i = i + 1
      sleep(2)
    rescue Interrupt
      break
    end
  end
end

run