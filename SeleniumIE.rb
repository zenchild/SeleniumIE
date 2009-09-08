=begin
-----------------------------------------------------------------------------
  Copyright © 2009 Dan Wanek <dwanek@nd.gov>
 
 
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

Original WatirMaker.rb license
-----------------------------------------------------------------------------
  This code is heavily inspired by the wonderful WatirMaker.rb program
  which was originally licensed under the following:

  license
  ---------------------------------------------------------------------------
  Copyright (c) 2004-2005, Michael S. Kelly, John Hann, and Scott Hanselman
  All rights reserved.
  
  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  
  1. Redistributions of source code must retain the above copyright notice,
  this list of conditions and the following disclaimer.
  
  2. Redistributions in binary form must reproduce the above copyright
  notice, this list of conditions and the following disclaimer in the
  documentation and/or other materials provided with the distribution.
  
  3. Neither the names Scott Hanselman, Michael S. Kelly nor the names of 
  contributors to this software may be used to endorse or promote products 
  derived from this software without specific prior written permission.
  
  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ``AS
  IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
  THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
  PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR
  CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
  EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
  PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
  OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
  WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
  OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
  ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
  --------------------------------------------------------------------------
=end

# This program attempts to output Selenium code within RSpec tests that is generated
# by attaching to a Win32ole Internet Explorer application and listening for events.
# I hope that this will provide similar functionality for IE that the wonderful
# SeleniumIDE for Firefox already achieves.
require 'win32ole'


class SeleniumIERecorder

	# ========================== Public Methods ========================= 
	public

	@@top_level_frame_name = ""

	def initialize
		@browser = WIN32OLE.new( 'InternetExplorer.Application' )
		@browser.visible = true

		@debug = File.new("selenium_debug.txt", 'w+') if $DEBUG
		@outfile = File.new("selenium_out.txt", 'w+')

		# contains references to all document objects that we're listening to
		@activeDocuments = Hash.new

		# This variable is checked to see if a statement should be written to the output file
		@record = true
		
		# stores the last frame name      
		@lastFrameName = ""

		setDebugLevel(2)
		
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
		if targetFrameName == @@top_level_frame_name
			@frameNames = Array.new
		end
		@frameNames << targetFrameName 
	end

	# http://msdn.microsoft.com/en-us/library/aa768329(VS.85).aspx
	def DocumentComplete(pDisp, url)
		begin
			# Check to see if document is loaded and catch the exception if it occurs.
			document = pDisp.document

			@debug.puts "************** Document Complete: #{url}" if $DEBUG
			@debug.puts "LocationURL: #{@browser.LocationURL}" if $DEBUG
			@debug.puts "URL: #{url}" if $DEBUG
			if( @record and @browser.LocationURL == url ) then
				@outfile.puts seleniumStatement("@browser.open(\"#{url}\")")
				@outfile.puts seleniumStatement("@browser.wait_for_page_to_load(30000)")
				@record = false
			end
			

			@frameNames.each do |frameName|
				if frameName == @@top_level_frame_name then
					@debug.puts "TOP: #{frameName}" if $DEBUG
					#document = pDisp.document
				else
					@debug.puts "FRAME: #{frameName}" if $DEBUG
				end
				
				documentKey = frameName
				if ( pDisp.Type == "HTML Document"  && !@activeDocuments.has_key?( documentKey ) ) then
					# create a new document object in the hash
					@activeDocuments[documentKey] = WIN32OLE_EVENT.new( document, 'HTMLDocumentEvents2' )

					# register event handlers
					@debug.puts "Adding document handler for '#{documentKey}'" if $DEBUG
					@activeDocuments[documentKey].on_event( 'onclick' ) { |*args| document_onclick( args[0] ) }
					@activeDocuments[documentKey].on_event( 'onfocusout' ) { |*args| document_onfocusout( args[0] ) }
					@activeDocuments[documentKey].on_event( 'onsubmit' ) { |*args| document_onsubmit( args[0] ) }
				end
			end
 
		rescue WIN32OLERuntimeError => e
			if e.to_s.match( "nknown property or method `document'" )
				@debug.puts "Document not yet loaded" if $DEBUG
				return
			else
				raise e
			end
		end
	end

	# http://msdn.microsoft.com/en-us/library/aa768334(VS.85).aspx
=begin
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
=end

	# http://msdn.microsoft.com/en-us/library/cc136549(VS.85).aspx
	def OnQuit(*ev_args)
		@browser.visible = false
		throw :done
	end

	# http://msdn.microsoft.com/en-us/library/aa768347(VS.85).aspx
	#def ProgressChange(nProgress, nProgressMax)
	#	@debug.puts "Progress Changed: #{nProgress}" if $DEBUG
	#end

	# http://msdn.microsoft.com/en-us/library/aa768348(VS.85).aspx
	#def PropertyChange(sProperty)
	#	@debug.puts "Property Changed: #{sProperty}" if $DEBUG
	#end

	# http://msdn.microsoft.com/en-us/library/aa768349(VS.85).aspx
	#def StatusTextChange(sText)
		#@debug.puts "Status Text Changed: #{sText}" if $DEBUG
	#end

	# http://msdn.microsoft.com/en-us/library/aa768350(VS.85).aspx
	#def TitleChange(sText)
	#	@debug.puts "Title changed to #{sText}" if $DEBUG
	#end

	# http://msdn.microsoft.com/en-us/library/aa768358(VS.85).aspx
	#def WindowStateChanged(dwFlags, dwValidFlagsMask)
	#	@debug.puts "Window State Changed" if $DEBUG
	#end
	
	def document_onsubmit (eventObj)
		printDebugComment( "Submitting Form:  " + eventObj.srcElement.tagName)
	end
	
	def get_form_input (form)
		form.all.length.times do |i|
			elem = form.all.item(i)
			if elem.tagName == "INPUT"
				printDebugComment( "Getting Form Input Tag/Type: #{elem.tagName}, #{elem.Type} " )
				if elem.Type == "text" or elem.Type == "password"
					@outfile.puts seleniumStatement( "@browser.type \"#{ getXpath(elem) }\", \"#{elem.Value}\"" )
				end
			end
		end
	end

	# ---------------------- methods from WatirMaker ---------------
	#
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Handles document onclick events.
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def document_onclick( eventObj )
		# if the user clicked something and the URL chandes as a result, it's probably due to this...
		
		case eventObj.srcElement.tagName
		when "INPUT", "A", "SPAN", "IMG", "TD"
			#@outfile.puts writeWatirStatement( eventObj, "click" )           
			get_form_input (eventObj.srcElement.form)

			@outfile.puts seleniumStatement( "@browser.click \"#{getXpath(eventObj.srcElement)}\", :wait_for => :page" )           
		else
			how, what = getHowWhat( eventObj.srcElement )
			printDebugComment( "Unsupported onclick tagname " + eventObj.srcElement.tagName +
							  " (" + how + ", '" + what + "')" )
		end
	end

	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Print Watir statement for onfocusout events.
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def document_onfocusout( eventObj ) 
		case eventObj.srcElement.tagName
		when "INPUT"
			case eventObj.srcElement.getAttribute( "type" )
			when "text", "password"
				printDebugComment("In OnFocusOut: TEXT/PASS: #{eventObj.srcElement.tagName}")
				@outfile.puts writeWatirStatement( eventObj, "set", eventObj.srcElement.value )
			when "checkbox"
				printDebugComment("In OnFocusOut: CHECKBOX")
				if eventObj.srcElement.checked
					@outfile.puts writeWatirStatement( eventObj, "set" )
				else
					@outfile.puts writeWatirStatement( eventObj, "clear" )
				end
			else
				printDebugComment( "Unsupported onfocusout INPUT type " + \
								  eventObj.srcElement.getAttribute( "type" ))
			end
		when "SELECT"
			@outfile.puts writeWatirStatement( eventObj, "select", eventObj.srcElement.value )
		else
			how, what = getHowWhat( eventObj.srcElement )
			printDebugComment( "Unsupported onfocusout tagname " + eventObj.srcElement.tagName + " (" + how + ", '" + what + "')" )
		end
	end


	def get_tag_with_id (element)
		return("#{element.tagName.downcase}[@id='#{element.Id}']")
	end

	def get_tag_with_name (element)
		return("#{element.tagName.downcase}[@name='#{element.Name}']")
	end

	def get_tag (element)
		return(element.tagName.downcase)
	end
	
	def getXpath (element)
		xpath = []

		i = 0
		# Set the max times to loop.  This should be plenty to get a unique XPATH object
		maxloop = 4
		while i < maxloop and element != nil
			case element.ole_obj_help.to_s
			when "DispHTMLDivElement", "DispHTMLFormElement", "DispHTMLAnchorElement", "DispHTMLInputElement", "DispHTMLTable"
				xpath.unshift(get_tag_with_id(element))
			when "DispHTMLTableRow"
				xpath.unshift("#{get_tag(element)}[#{element.rowIndex+1}]")
			when "DispHTMLTableCell"
				xpath.unshift("#{get_tag(element)}[#{element.cellIndex+1}]")
			when "DispHTMLLIElement"
				xpath.unshift("#{get_tag(element)}[#{element.value}]")
			when "DispHTMLHtmlElement", "DispHTMLBody", "DispHTMLHeaderElement", "DispHTMLTableSection", "DispHTMLOListElement"
				xpath.unshift(get_tag(element))
			else
				printDebugComment("Unhandled XPATH element Type: #{element.ole_obj_help}")
				xpath.unshift("*")
			end

			i+=1
			element = element.parentNode
		end

		return "//" + xpath.join("/")
	end


	def getHowWhat( element )
		if element.getAttribute( "id" ) != nil && element.getAttribute( "id" ) != ""
			return ":id", element.getAttribute( "id" )
		elsif element.getAttribute( "name" ) != nil && element.getAttribute( "name" ) != ""
			return ":name", element.getAttribute( "name" )
		else
			#printDebugComment(element.tagName)
			case element.tagName
			when "A"
				return ":text", %Q{#{element.innerText}}
			else
				#return the index of this element in the 'all' collection as a string
				index = element.sourceIndex
				if index != nil
					return ":index", %Q{#{index}}
				end
			end
		end
	end


	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Generates a line of WATIR script based on the HTML element and the action to take.
	##
	## element: The IE HTML element to perform the action on.
	## action:  The WATIR action to perform on said element.
	## value:   The value to assign to the element (required for 'set' and 'select' actions)
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def writeWatirStatement( eventObj, action, value = "" )
		str = ""
		element = eventObj.srcElement
		if element == nil
			printDebugComment "writeWatirStatement eventObj.srcElement was nil!"
			return str
		end
		
		# if we're trasitioning between frames, insert a delay statement
		if @lastFrameName != element.document.parentWindow.name
			# TODO: How do we do a Thread.Wait() in Ruby?
		end
		
		case element.tagName
		when "INPUT"
			case element.getAttribute( "type" )
			when "submit", "image", "button"
				if action == "click"
					str = genWatirAccessor( "button", element ) + action
				end
			when "text", "password"
				if action == "set"
					str = genWatirAccessor( "text_field", element ) + action + "( '" +  element.value + "' )"
				end
			when "checkbox"
				if action == "set" || action == "clear"
					str = genWatirAccessor( "checkbox", element ) + action
				end
			when "radio"
				if action == "set" || action == "clear"
					str = genWatirAccessor( "radio", element ) + action
				end
			else
				how, what = getHowWhat( eventObj.srcElement )
				printDebugComment( "Unsupported INPUT type " + element.getAttribute( "type" ) +
								  " (" + how + ", '" + what + "')" )
			end
		when "A"
			if action == "click"
				str = genWatirAccessor( "link", element ) + action
			end
		when "SPAN"
			if action == "click"
				str = genWatirAccessor( "span", element ) + action
			end
		when "IMG"
			if action == "click"
				str = genWatirAccessor( "image", element ) + action
			end
		when "TD"
			if action == "click"
				how, what = getHowWhat( element )
				str = genIePrefix( element ) + "document.all[ '" + what + "' ].click"
			end
		when "SELECT"
			if action == "select"
				for i in 0..element.options.length-1
					if element.options[ %Q{#{i}} ].selected
						str += genWatirAccessor( "select_list", element ) + action + "( '" + \
						element.options[ %Q{#{i}} ].text + "' )\n"
					end
				end
			end
		else
			how, what = getHowWhat( eventObj.srcElement )
			printDebugComment( "Unsupported onclick tagname " + eventObj.srcElement.tagName +
							  " (" + how + ", '" + what + "')" )
		end
		
		if str != ""
			@lastFrameName = element.document.parentWindow.name
		else
			printDebugComment "Unsupported action '" + action + "' for '" + element.tagName + "'."
		end
		return str
	end
 
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Generates the WATIR code necessary for accessing a particular document element.
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def genWatirAccessor( watirType, element )
		iePrefix = genIePrefix( element )
		how, what = getHowWhat( element )
		
		# for some reason the index 'How' doesn't work for the index we get from our code
		if how == ":index"
			return iePrefix + "document.all[ '" + what + "' ]."
		elsif how == ":id"
			if what.include? "_"
				what = what[what.rindex("_") + 1,what.length - what.rindex("_")]
				what = "/" + what + "$/"
			else
				what = "'" + what + "'"
			end
			return iePrefix + watirType + "( " + how + ", " + what + " )."
		else
			return iePrefix + watirType + "( " + how + ", '" + what + "' )."
		end
	end
 
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Generates the ie prefix necessary for accessing a particular document element, including frames.
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def genIePrefix( element )
		printDebugComment("genIePrefix")
		parentWindowName = element.document.parentWindow.name
		if parentWindowName != @@top_level_frame_name
			return "@browser.frame( :name, '" + parentWindowName + "' )."
		else
			return "@browser."
		end
	end

	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Print warning comment.
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def printWarningComment( warning )
		if @printWarnings
			puts ""
			puts "# WARNING: '" + warning
			puts ""
		end
	end


	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Print a debug message comment.
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def printDebugComment( message )
		if @printDebugInfo
			puts "# DEBUG: " + message
		end
	end


	##//////////////////////////////////////////////////////////////////////////////////////////////////
	##
	## Print comment showing the page navigated to.
	##
	##//////////////////////////////////////////////////////////////////////////////////////////////////
	def printNavigateComment( url, frameName )
		puts ""
		puts "# frame loading: '" + frameName + "'" if frameName != nil
		puts "# navigating to: '" + url + "'"
		puts ""
	end
	
	def setDebugLevel(level)
		if level >= 2
			@printDebugInfo = true
		else
			@printDebugInfo = false			
		end
		if level >= 1
			@printWarnings = true
		else
			@printWarnings = false
		end
   end
end  # class SeleniumIERecorder


recorder = SeleniumIERecorder.new
recorder.record
