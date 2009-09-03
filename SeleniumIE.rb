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
		
		events = WIN32OLE_EVENT.new( @browser, 'DWebBrowserEvents2' )
		events.on_event { |*ev_args| eventHandler( *ev_args ) }
	end


	# This is the main loop that listens for events coming in from IE and dispatches them.
	def record
		trap('INT') { puts ""; throw :done }
		
		puts header
		catch(:done) do
			loop do
				WIN32OLE_EVENT.message_loop
			end
		end
		puts footer
	end


	# ========================= Private Methods ========================= 
	private

	# eventHandler is the dispatcher for all incoming events.
	def eventHandler(ev_name, *ev_args)
		if methodExists?(ev_name)
			method(ev_name).call(*ev_args)
			puts "-------------------------------------------------" if $DEBUG
		end
	end

	# This writes out the Selenium/RDspec statement with an optional code indentation argument.
	def seleniumStatement(statment, indent=0)
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
		puts "Navigating to #{url}"
		puts "\t Target Frame: #{targetFrameName}"
	end

	# http://msdn.microsoft.com/en-us/library/aa768329(VS.85).aspx
	def DocumentComplete(pDisp, url)
	end

	# http://msdn.microsoft.com/en-us/library/aa768334(VS.85).aspx
	def NavigateComplete2(pDisp, url)
		puts "** Navigation to #{url} complete **"
	end

	# http://msdn.microsoft.com/en-us/library/bb268221(VS.85).aspx
	def NavigateError(pDisp, url, targetFrameName, statusCode, cancel)
		puts "Navigation ERROR occured (#{url}):  Status Code #{statusCode}"
	end

	# http://msdn.microsoft.com/en-us/library/aa768336(VS.85).aspx
	def NewWindow2(ppDisp, cancel)
		puts "NewWindow2 event fired"
	end

	# http://msdn.microsoft.com/en-us/library/aa768337(VS.85).aspx
	def NewWindow3(ppDisp, cancel, dwFlags, bstrUrlContext, bstrUrl)
		puts "NewWindow3 event fired"
	end

	# http://msdn.microsoft.com/en-us/library/cc136549(VS.85).aspx
	def OnQuit(*ev_args)
		@browser.visible = false
		throw :done
	end

	# http://msdn.microsoft.com/en-us/library/aa768347(VS.85).aspx
	def ProgressChange(nProgress, nProgressMax)
	end

	# http://msdn.microsoft.com/en-us/library/aa768348(VS.85).aspx
	def PropertyChange(sProperty)
		puts "Property Changed: #{sProperty}"
	end

	# http://msdn.microsoft.com/en-us/library/aa768349(VS.85).aspx
	def StatusTextChange(sText)
		puts "Status Text Changed: #{sText}"
	end

	# http://msdn.microsoft.com/en-us/library/aa768350(VS.85).aspx
	def TitleChange(sText)
		puts "Title changed to #{sText}"
	end

	# http://msdn.microsoft.com/en-us/library/aa768358(VS.85).aspx
	def WindowStateChanged(dwFlags, dwValidFlagsMask)
	end

end


recorder = SeleniumIERecorder.new
recorder.record
