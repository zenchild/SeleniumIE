# This program attempts to output Selenium code within RSpec tests that is generated
# by attaching to a Win32ole Internet Explorer application and listening for events.
# I hope that this will provide similar functionality for IE that the wonderful
# SeleniumIDE for Firefox already achieves.
require 'win32ole'


class SeleniumIERecorder

	# ========================== Public Methods ========================= 
	public

	def initialize
		@browser = WIN32OLE.new( 'InternetExplorer.Application' )
		@browser.visible = true

		@debug = File.new("selenium_debug.txt", 'w+') if $DEBUG
		@outfile = File.new("selenium_out.txt", 'w+')

		# This variable is checked to see if a statement should be written to the output file
		@record = true
		
		events = WIN32OLE_EVENT.new( @browser, 'DWebBrowserEvents2' )
		events.on_event { |*ev_args| eventHandler( *ev_args ) }
	end


	# This is the main loop that listens for events coming in from IE and dispatches them.
	def record
		trap('INT') { puts ""; throw :done }
		
		@outfile.puts header
		catch(:done) do
			loop do
				WIN32OLE_EVENT.message_loop
			end
		end
		@outfile.puts footer
	end


	# ========================= Private Methods ========================= 
	private

	# eventHandler is the dispatcher for all incoming events.
	def eventHandler(ev_name, *ev_args)
		if methodExists?(ev_name)
			method(ev_name).call(*ev_args)
			@debug.puts "-------------------------------------------------" if $DEBUG
		end
	end

	# This writes out the Selenium/RDspec statement with an optional code indentation argument.
	def seleniumStatement(statement, indent=1)
		return(<<-EOS.gsub(/^\t*/, "\t" * indent))
			#{statement}
		EOS
	end


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

	def methodExists?(method_name)
		return self.methods.include?(method_name)
	end

	# ========================= IE Event Methods ========================
	# These need to be public if the methodsExists? method is to work
	public

	# http://msdn.microsoft.com/en-us/library/aa768280(VS.85).aspx
	def BeforeNavigate2(pDisp, url, flags, targetFrameName, postData, headers, cancel)
		@debug.puts "Navigating to #{url}" if $DEBUG
		@debug.puts "\t Target Frame: #{targetFrameName}" if $DEBUG
		@debug.puts "\t Location URL: #{@browser.LocationURL}" if $DEBUG
		if @browser.LocationURL == "" and @record then
			@outfile.puts seleniumStatement("@browser.open(#{url})")
			@record = false
		end
	end

	# http://msdn.microsoft.com/en-us/library/aa768329(VS.85).aspx
	def DocumentComplete(pDisp, url)
		@debug.puts "Document Complete: #{url}" if $DEBUG
		if url == @browser.LocationURL
			@record = true
			# register each frame with an event handler
			# @browser.Document.frames.length.times do |i|
			# 	register @browser.Document.frames.item(i)
			# end
		end
	end

	# http://msdn.microsoft.com/en-us/library/aa768334(VS.85).aspx
	def NavigateComplete2(pDisp, url)
		@debug.puts "** Navigation to #{url} complete **" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/bb268221(VS.85).aspx
	def NavigateError(pDisp, url, targetFrameName, statusCode, cancel)
		@debug.puts "Navigation ERROR occured (#{url}):  Status Code #{statusCode}" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/aa768336(VS.85).aspx
	def NewWindow2(ppDisp, cancel)
		@debug.puts "NewWindow2 event fired" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/aa768337(VS.85).aspx
	def NewWindow3(ppDisp, cancel, dwFlags, bstrUrlContext, bstrUrl)
		@debug.puts "NewWindow3 event fired" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/cc136549(VS.85).aspx
	def OnQuit(*ev_args)
		@browser.visible = false
		throw :done
	end

	# http://msdn.microsoft.com/en-us/library/aa768347(VS.85).aspx
	def ProgressChange(nProgress, nProgressMax)
		@debug.puts "Progress Changed: #{nProgress}" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/aa768348(VS.85).aspx
	def PropertyChange(sProperty)
		@debug.puts "Property Changed: #{sProperty}" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/aa768349(VS.85).aspx
	def StatusTextChange(sText)
		@debug.puts "Status Text Changed: #{sText}" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/aa768350(VS.85).aspx
	def TitleChange(sText)
		@debug.puts "Title changed to #{sText}" if $DEBUG
	end

	# http://msdn.microsoft.com/en-us/library/aa768358(VS.85).aspx
	def WindowStateChanged(dwFlags, dwValidFlagsMask)
		@debug.puts "Window State Changed" if $DEBUG
	end

end


recorder = SeleniumIERecorder.new
recorder.record
