=begin
-----------------------------------------------------------------------------
  Copyright � 2009 Dan Wanek <dwanek@nd.gov>


  This file is part of SeleniumIE.

  SeleniumIE is free software: you can redistribute it and/or
  modify it under the terms of the GNU General Public License as published
  by the Free Software Foundation, either version 3 of the License, or (at
  your option) any later version.

  SeleniumIE is distributed in the hope that it will be useful,
  but WITHOUT ANY WARRANTY; without even the implied warranty of
  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General
  Public License for more details.

  You should have received a copy of the GNU General Public License along
  with SeleniumIE.  If not, see <http://www.gnu.org/licenses/>.
-----------------------------------------------------------------------------
=end

# This program attempts to output Selenium code within RSpec tests that is generated
# by attaching to a Win32ole Internet Explorer application and listening for events.
# I hope that this will provide similar functionality for IE that the wonderful
# SeleniumIDE for Firefox already achieves.
#
# I would like to thank Michael S. Kelly, John Hann, and Scott Hanselman for their
# work on WatirMaker.rb, which inspired this code.  Early versions of this script
# contained quite a bit of the code used in WatirMaker.rb.
require 'win32ole'


class SeleniumIERecorder

	# ========================== Public Methods =========================
	public

	@@top_level_frame_name = ""

	def initialize
		@debugfile = File.new("selenium_debug.txt", 'w+')
		@outfile = File.new("selenium_out.txt", 'w+')

		# contains references to all document objects that we're listening to
		@activeDocuments = Hash.new

		# stores the string for the last onclick URL so we can do duplicate detection
		@last_onclick = ""

		@mouse_clicked = false

		# Keep track of what RSPEC test case we are on
		@test_case = 1

		# This is set by the WindowStateChange event to determine if we should add an "open" statement or
		# just use the mouse click navigation.  This event doesn't directly coorespond to a manual URL entry
		# but for most intents and purposes of recording a transaction it should work fine.
		@navigate_directly = true

		@browser = WIN32OLE.new( 'InternetExplorer.Application' )
		@browser.visible = true

		@events = WIN32OLE_EVENT.new( @browser, 'DWebBrowserEvents2' )
		@events.on_event { |*ev_args| eventHandler( *ev_args ) }
	end


	# This is the main loop that listens for events coming in from IE and dispatches them.
	def record
		trap('INT') { puts ""; throw :done }

		@outfile.puts header
		@outfile.puts rspec_begin
		catch(:done) do
			loop do
				WIN32OLE_EVENT.message_loop
			end
		end
		@outfile.puts rspec_end
		@outfile.puts footer
	end


	# ========================= Private Methods =========================
	private

	# eventHandler is the dispatcher for all incoming events.
	def eventHandler(ev_name, *ev_args)
		#debug_msg "------------------- #{ev_name} ---------------------"
		if methodExists?(ev_name)
			method(ev_name).call(*ev_args)
		end
	end

	# This writes out the Selenium/RDspec statement with an optional code indentation argument.
	def selenium_statement(statement, indent=2)
		debug_msg "----------------------------------------------"
		debug_msg "REC: #{statement}"
		debug_msg "----------------------------------------------"
		return(<<-EOS.gsub(/^\t*/, "\t" * indent))
			#{statement}
		EOS
	end

	def rspec_begin(descriptor="this is test case #{@test_case}")
		return selenium_statement("it \"descriptor\" do",1)
	end

	def rspec_end()
		return selenium_statement("end",1)
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
				before(:all) do
					@browser = Selenium::Client::Driver.new("localhost", 4444, "*iexplore", "http://", 10000);
					@browser.start
				end

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
		if targetFrameName == @@top_level_frame_name
			@frameNames = Array.new
		end
		@frameNames << targetFrameName
	end

	# http://msdn.microsoft.com/en-us/library/aa768329(VS.85).aspx
	def DocumentComplete(pDisp, url)
		begin
			# Check to see if document is loaded and catch the exception if it occurs.
			# document = pDisp.document
			# Check to make sure the @browser object is ready for use
			if( @browser.ReadyState != 4 )
				debug_msg("Browser not in \"complete\" ReadyState")
				return
			end

			debug_msg "************** Document Complete: #{url}"
			debug_msg "LocationURL: #{@browser.LocationURL}"
			debug_msg "URL: #{url}"
			if( @navigate_directly ) then
				@outfile.puts selenium_statement("@browser.open(\"#{url}\")")
				@outfile.puts selenium_statement("@browser.wait_for_page_to_load(30000)")
				@navigate_directly = false
			end


			@frameNames.each do |frameName|
				if frameName == @@top_level_frame_name then
					debug_msg "TOP: #{frameName}"
					document = pDisp.document
				else
					debug_msg "FRAME: #{frameName}"
					document = pDisp.document.frames[frameName].document
				end

				forms = document.forms
				forms.length.times do |i|
					forms.item(i).onsubmit = method("form_onsubmit")
				end

				document.onclick = ""

				documentKey = frameName
				#if ( pDisp.Type == "HTML Document"  && !@activeDocuments.has_key?( documentKey ) ) then
				if ( pDisp.Type == "HTML Document" )
					# create a new document object in the hash
					if( @activeDocuments.key?(documentKey) ) then
						@activeDocuments[documentKey] = nil
					end

					@activeDocuments[documentKey] = WIN32OLE_EVENT.new( document, 'HTMLDocumentEvents2' )

					# register event handlers
					debug_msg "Adding document handler for '#{documentKey}'"
					@activeDocuments[documentKey].on_event( 'onclick' ) { |*args| document_onclick( args[0] ) }
				end
			end

		rescue WIN32OLERuntimeError => e
			if e.to_s.match( "nknown property or method" )
				debug_msg "Document not yet loaded"
				return
			elsif e.to_s.match( "Access is denied" )
				debug_msg "Method access denied"
				return
			else
				raise e
			end
		end
	end

	def form_onsubmit()
		form = @browser.Document.activeElement.form
		#puts "***** ACCCCCCCCCCCCT #{ @browser.Document.activeElement.tagName }"
		#if @browser.Document.activeElement.tagName == "INPUT"
		#	puts "***** ACCCCCCCCCCCCT #{ @browser.Document.activeElement.Type }"
		#end
		#puts "***** ACCCCCCCCCCCCT #{ @browser.Document.activeElement.Id }"
		get_form_input( form )
		@outfile.puts selenium_statement( "@browser.submit \"#{ getXpath(form) }\"" ) unless @mouse_clicked
	end

	# http://msdn.microsoft.com/en-us/library/aa768334(VS.85).aspx
=begin
	def NavigateComplete2(pDisp, url)
		debug_msg "** Navigation to #{url} complete **"
	end

	# http://msdn.microsoft.com/en-us/library/bb268221(VS.85).aspx
	def NavigateError(pDisp, url, targetFrameName, statusCode, cancel)
		debug_msg "Navigation ERROR occured (#{url}):  Status Code #{statusCode}"
	end

	# http://msdn.microsoft.com/en-us/library/aa768336(VS.85).aspx
	def NewWindow2(ppDisp, cancel)
		debug_msg "NewWindow2 event fired"
	end

	# http://msdn.microsoft.com/en-us/library/aa768337(VS.85).aspx
	def NewWindow3(ppDisp, cancel, dwFlags, bstrUrlContext, bstrUrl)
		debug_msg "NewWindow3 event fired"
	end
=end

	# http://msdn.microsoft.com/en-us/library/cc136549(VS.85).aspx
	def OnQuit(*ev_args)
		@browser.visible = false
		throw :done
	end

	# http://msdn.microsoft.com/en-us/library/aa768347(VS.85).aspx
	#def ProgressChange(nProgress, nProgressMax)
	#	debug_msg "Progress Changed: #{nProgress}"
	#end

	# http://msdn.microsoft.com/en-us/library/aa768348(VS.85).aspx
=begin
	def PropertyChange(sProperty)
		debug_msg "Property Changed: #{sProperty}"
		prop = @browser.GetProperty(sProperty)
		if prop != nil
			debug_msg "\tPROP: #{prop.class.to_s}"
			debug_msg "\tPROP: #{prop}"
			debug_msg "\tPROP: #{str}"
		end
	end
=end

	# http://msdn.microsoft.com/en-us/library/aa768349(VS.85).aspx
	#def StatusTextChange(sText)
		#debug_msg "Status Text Changed: #{sText}"
	#end

	# http://msdn.microsoft.com/en-us/library/aa768350(VS.85).aspx
	#def TitleChange(sText)
	#	debug_msg "Title changed to #{sText}"
	#end

	# http://msdn.microsoft.com/en-us/library/aa768358(VS.85).aspx
	def WindowStateChanged(dwFlags, dwValidFlagsMask)
		debug_msg( "Window State Changed: FLAGS: #{dwFlags}  MASK: #{dwValidFlagsMask}" )
		if dwFlags == 3 and dwValidFlagsMask == 3 then
			@navigate_directly = true
		end
	end

	def get_form_input (form)
		form.all.length.times do |i|
			elem = form.all.item(i)
			if elem.tagName == "INPUT"
				debug_msg( "Getting Form Input Tag/Type: #{elem.tagName}, #{elem.Type} " )
				if elem.Type == "text" or elem.Type == "password"
					@outfile.puts selenium_statement( "@browser.type \"#{ getXpath(elem) }\", \"#{elem.Value}\"" )
				end
			end
		end
	end

	def get_tag_selector(element)
		elem_tag = element.tagName.downcase
		case elem_tag
		when "tr"
			return( "#{elem_tag}" + "[#{element.sectionRowIndex+1}]" )
		when "td"
			return( "#{elem_tag}" + "[#{element.cellIndex+1}]" )
		when "li"
			return( "#{elem_tag}" + "[#{element.value+1}]" )
		else
			if ((elem_id = element.getAttribute('id')) != nil) then
				return( "#{elem_tag}" + "[@id='#{elem_id}']" )
			elsif ((elem_name = element.getAttribute('name')) != nil) then
				return( "#{elem_tag}" + "[@name='#{elem_name}']" )
			elsif ((elem_href = element.getAttribute('href')) != nil) then
				return( "#{elem_tag}" + "[@href='#{elem_href}']" )
			else
				return(element.tagName.downcase)
			end
		end
	end

	def getXpath (element)
		xpath = []

		i = 0
		# Set the max times to loop.  This should be plenty to get a unique XPATH object
		maxloop = 10
		while i < maxloop and element != nil
			case element.ole_obj_help.to_s
			when "DispHTMLDivElement", "DispHTMLFormElement", "DispHTMLInputElement", "DispHTMLButtonElement",
				"DispHTMLTable", "DispHTMLTableRow", "DispHTMLTableCell", "DispHTMLLIElement", "DispHTMLHtmlElement",
				"DispHTMLBody", "DispHTMLHeaderElement", "DispHTMLTableSection", "DispHTMLFontElement", "DispHTMLUListElement",
				"DispHTMLOListElement", "DispHTMLParaElement", "DispHTMLPhraseElement", "DispHTMLAnchorElement",
				"DispHTMLSpanElement"
				xpath.unshift(get_tag_selector(element))
			when "DispHTMLDocument"
				# End of the road
				i = maxloop
			else
				debug_msg("Unhandled XPATH element Type: #{element.ole_obj_help}  TAG: #{element.tagName} ")
				if(element.getAttribute('tagName') != nil) then
					xpath.unshift(get_tag_selector(element))
				else
					xpath.unshift("*")
				end
			end

			i+=1
			element = element.parentNode
		end


		return '//' + xpath.join('/')
	end


	# This method is a sink for Document (IHTMLDocument2) onclick events.
	# http://msdn.microsoft.com/en-us/library/aa752574(VS.85).aspx
	def document_onclick( eventObj )
		debug_msg "ONCLICK: #{eventObj.srcElement.getAttribute('tagName')}"
		str = ""

		case eventObj.srcElement.tagName
		when "INPUT", "A"
			case eventObj.srcElement.getAttribute('Type')
			when "text", "password", "hidden"
				# don't register onclick for these types
			when "button", "checkbox", "file", "image", "radio", "reset", "submit"
				if eventObj.srcElement.getAttribute('form') != nil then
					get_form_input(eventObj.srcElement.getAttribute('form'))
				end
			end
			str = "@browser.click \"#{getXpath(eventObj.srcElement)}\", :wait_for => :page"
			if str == @last_onclick
				str = "DUP? #{str}"
			end
			@outfile.puts selenium_statement( str )
		when "A"
			str = "@browser.click \"#{getXpath(eventObj.srcElement)}\", :wait_for => :page"
			if str == @last_onclick
				str = "DUP? #{str}"
			end
			@outfile.puts selenium_statement( str )
		when "BUTTON", "SPAN", "IMG", "TD"
			if eventObj.srcElement.getAttribute('form') == nil
				str = "@browser.click \"#{getXpath(eventObj.srcElement)}\""
			else
				str = "@browser.click \"#{getXpath(eventObj.srcElement)}\", :wait_for => :page"
			end

			if str == @last_onclick
				str = "DUP? #{str}"
			end
			@outfile.puts selenium_statement( str )
		else
			debug_msg( "Unsupported onclick tagname " + eventObj.srcElement.tagName )
		end

		@last_onclick = str
	end

	def debug_msg( message )
		if $DEBUG
			#@debugfile.puts "# DEBUG: " + message
			puts "# DEBUG: " + message
		end
	end

end  # class SeleniumIERecorder


recorder = SeleniumIERecorder.new
recorder.record
