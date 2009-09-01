# This program attempts to output Selenium code within RSpec tests that is generated
# by attaching to a Win32ole Internet Explorer application and listening for events.
# I hope that this will provide similar functionality for IE that the wonderful
# SeleniumIDE for Firefox already achieves.
require 'win32ole'


module BrowserEvents
	def methodExists?(method_name)
		return self.methods.include?(method_name)
	end

	def BeforeNavigate2(*ev_args)
		puts "BeforeNavigate2: #{ev_args.class.to_s}"
	end

	def OnQuit(*ev_args)
		throw :done
	end
end



class SeleniumIERecorder
	include BrowserEvents

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
			method(ev_name).call(ev_args)
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
end


recorder = SeleniumIERecorder.new
recorder.record
