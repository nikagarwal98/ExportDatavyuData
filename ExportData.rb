# The Datavyu scripting API provides a scriptable interface to Datavyu’s spreadsheet, 
# allowing you to manipulate your data, export in any format you’d like, or check your data for errors.
require 'Datavyu_API.rb'

# All the code should go under this block
begin

  # This folder contains all the files whose data we want to extract
  filedir = File.expand_path("C:/Users/Nikunj/OneDrive - Grinnell College/Grinnell/Emma Kelty Stephen Research/datafiles/") + "/"

  # Build an array of the list of files of the folder
  filenames = Dir.new(filedir).entries

  # Iterate over the array of files
  for file in filenames

    #Check if the file is of .opf type
    if file.include?(".opf")

      # Print to console if the file is successfully loaded
      puts "LOADING DATABASE: " + filedir+file
      $db,proj = load_db(filedir+file)
      puts "SUCCESSFULLY LOADED"

      # Get column variables from Datavyu and save it to local variables
      id = getVariable("id")
      trial = getVariable("trial")
      section = getVariable("section")
      look = getVariable("look")
      directions = ["left", "right", "center", "away"]
      participant = id.cells[0]  

      # Header [1st row of the csv]
      header = ["Subject Number", "Birthdate", "Sex", "Testdate", "Counterbalance"] 
      16.times do |i|
        10.times do
          header << "Trial " + (i+1).to_s
        end
      end

      # Header [2nd row of the csv]
      subheader = []
      5.times do
        subheader << ""
      end
      16.times do
        5.times do
          subheader << "Non-directive Audio"
        end
        5.times do
          subheader << "Directive Audio"
        end
      end

      # Header [3rd row of the csv]
      subheader2 = []
      5.times do
        subheader2 << ""
      end
      32.times do
        subheader2 << directions
        subheader2 << "Switches"
      end

      # Trail Class for each trail; contains trail number and audiotype object
      class Trial
        attr_accessor :trialnum, :audio
        def initialize(trialnum, audio)
          @trialnum = trialnum
          @audio = audio
        end
      end

      # Audiotype class for each audiotype in each trial; contains audiotype (directive/non-directive) and an array of "Look" objects
      class Audiotype
        attr_accessor :audiotype, :entries
        def initialize(audiotype, entries)
          @audiotype = audiotype
          @entries = entries
        end
      end


      # Look class contains the start time, end time and the direction of the look
      class Look
        attr_accessor :timein, :timeout, :direction
        def initialize(timein, timeout, direction)
          @timein = timein
          @timeout = timeout
          @direction = direction
        end
      end

      # Method to return the direction of the look 
      def dir(direction)
        if direction == "f"
          return "left"
        elsif direction == "g"
          return "center"
        elsif direction == "h"
          return "right"
        elsif direction == "c"
          return "away"
        else
          return "invalid"
        end
      end

      # Method to return the audiotype of the audio
      def audiotype(audio)
        if audio == "n"
          return "Non-directive Audio"
        elsif audio == "d"
          return "Directive Audio"
        else
          return "invalid type"
        end 
      end

      # Method to iterate over indivual entries of an audiotype and return the total time a child has looked in a particular direction
      def timeInEachDirection(audiotype, direction)
        time = 0
        audiotype.each do |entry|
          if entry.direction == direction
            time = time + entry.timeout - entry.timein
          end
        end
        return time
      end

      # Method to iterate over each look in an audiotype and return the number of switches
      def switches(audio)
        count = 0
        audio.each do |look|
          count = count + 1
        end
        return count
      end

      arrTrial = []
      arrAudiotype = []
      arrLook = []

      # Iterate over each trial
      for tcell in trial.cells

        # Iterate over each section (audio) within the trial
        for scell in section.cells
          if scell.onset >= tcell.onset && scell.offset <= tcell.offset

            # Iterate over each look within the audio
            for lcell in look.cells
              if lcell.onset >= scell.onset && lcell.offset <= scell.offset

                # Create an array of looks within an audio
                arrLook << Look.new(lcell.onset, lcell.offset, dir(lcell.direction))

              end
            end

            # Create an array of the two audiotypes in a trial
            arrAudiotype << Audiotype.new(audiotype(scell.audiotype), arrLook)
            arrLook = []
          end
        end

        # Create an array of trials
        arrTrial << Trial.new(tcell.trialnumber, arrAudiotype)
        arrAudiotype = []
      end

      # Add data for participant
      participantData = [participant.subjectnumber, participant.birthdate, participant.sex, participant.testdate, participant.counterbalance]

      # Sort the data in the trial according to trial number
      arrTrial = arrTrial.sort_by {|data| data.trialnum.to_i}

      # Iterate over each trial
      arrTrial.each do |trial|

        # Iterate over each audiotype in the trial
        trial.audio.each do |y|

          # Iterate over the array of directions (left, right, center, away)
          directions.each do |dir|

            # Find the percantage of time the child looked in each direction
            participantData << ((timeInEachDirection(y.entries, dir)/60.to_f).round(2)).to_s + "%"
          end

          # Add the number of switches for each audiotype
          participantData << switches(y.entries)
        end
      end

      # Join all the data we have collected - header row 1, header row 2, header row 3 and participant data
      data = [header.join(","), subheader.join(","), subheader2.join(","), participantData.join(",")]
      puts "Writing to file..."

      # Set output file path
      out_file = File.expand_path("C:/Users/Nikunj/OneDrive - Grinnell College/Grinnell/Emma Kelty Stephen Research/datafiles/"+participant.subjectnumber+".csv")
      out = File.new(out_file,'w')

      # Output data
      out.puts data

      puts "FINISHED"
    
    end

  end

end
