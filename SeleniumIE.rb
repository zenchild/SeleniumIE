# This program attempts to output Selenium code within RSpec tests that is generated
# by attaching to a Win32ole Internet Explorer application and listening for events.
# I hope that this will provide similar functionality for IE that the wonderful
# SeleniumIDE for Firefox already achieves.


class SeleniumIERecorder

	# Write the header of the code output and include the basic setup.
	def header
		return(<<-EOS.gsub(/^\t{3}/,''))
			# =============================================================== #
			# Selenium RSPEC Script - Genererated from SeleniumIERecorder
			# Run with RSpec's "spec" command:
			# => spec <cmd_spec.rb>
			# =============================================================== #
			require 'selenium/client'
			require 'spec'

			describe "Test Case 1 (RENAME THIS)" do
		EOS
	end

	# Write the footer of the code output
	def footer
		return(<<-EOS.gsub(/^\t{3}/,''))
			end
			# =============================================================== #
		EOS
	end

	# This writes out the Selenium/RDspec statement with an optional code indentation argument.
	def seleniumStatement(statment, indent=0)
	end


	# This is the main loop that listens for events coming in from IE and dispatches them.
	def mainLoop
		trap('INT') { puts ""; throw :done }
		
		puts header
		catch(:done) do
			loop do
				puts "Hello there"
				sleep(10)
			end
		end
		puts footer
	end
end

t = Test.new
t.mainLoop
