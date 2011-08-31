$signature_survey_directory = "signature_survey" 
$root_directory_name = File.basename(Dir.pwd)
$file_extension = "vb"

$html_style_block = %Q% 
	<style type="text/css">
	    body { font-family: georgia, serif; }
	    h1 { color: #111111 }
	    li { font-family: Consolas, monospace; font-size: 14px; margin-bottom:20px; list-style-type: none; word-wrap:break-word; }
	    .line_count { color: #000080; font-family: Consolas, monospace;}
	    .method_count { color: #008080; font-family: Consolas, monospace;}
	    .signature { font-size: 12px; }
	    .inside_method { background-color: #eeeeee; }
	    .detail_page { font-family: Consolas, monospace; font-size: 14px; word-wrap: break-word}
	    .method_name { background-color: #ccccee ; font-weight: bold; }
	    .type { background-color: aaffbb; font-weight: bold; }
	    .inside_if { background-color:#CCEECC; }
	    .inside_loop { background-color:#EECCCC; }
	    .inside_loop .inside_if { background-color:#BBBBAA; }
	    .inside_if .inside_loop { background-color:#CCAAAA; }
	    .comment { color: #777777; }
	    p { display: inline; margin: 0px; padding: 0px; }
	    div { display: inline; margin: 0px; padding: 0px; }
	    #legend { font-family: Consolas, monospace; font-size: 14px; margin-bottom:10px; }
	    #legend td { padding-right: 15px; }
	</style>%

$legend_html = %Q%
    <table id="legend">
	<tr><td style="padding-bottom:10px;"><span style="font-size:18px; font-family: georgia, serif; font-weight: bold;">Legend</span></td></tr>
	<tr><td><span class="method_count">14m</span> means 14 methods</td> <td><span class="line_count">294L</span> means 294 lines</td> </tr>
	<tr><td>a space is a blank line</td><td>. is a (non-blank) line</td> </tr>
	<tr><td><span class='inside_if'>?</span> is a single line If statement</td><td>' is a commented line</td></tr>
        <tr><td><span class='inside_if'>|</span> is an Else or Case statement</td><td><div class='inside_method'>{Inside a method}</div></td></tr>
	<tr><td><span class='inside_if'>(Inside an if block)</span></td><td><p class='inside_loop'>[Inside a loop]</p></td></tr>
	<tr><td><p class='inside_loop'><span class='inside_if'>[(An if block inside a loop)]</span></p></td><td> <span class='inside_if'><p class='inside_loop'>([A loop inside an if block])</p></span></td></tr>
    </table>%

def for_each_vb_file_in_dir(directory, reportfile, method_to_execute)
    current_directory = Dir.new(directory)

    #don't want to start looking through the directories we are generating -- infinite recursion!!
    return if File.basename(current_directory.path) == $signature_survey_directory

    #make corresponding dir for sig survey files
    new_sig_survey_dir = current_directory.path
    new_sig_survey_dir[0] = $signature_survey_directory
    Dir.mkdir new_sig_survey_dir unless Dir.exists? new_sig_survey_dir

    #make title out of folder name
    puts current_directory.path
    reportfile.puts "<h3>#{current_directory.path}</h3>" 
    
    #process all vb files in current dir
    reportfile.puts "<ul>"
    Dir.glob(current_directory.path + "/*.#{$file_extension}") { |filename|
	puts "Processed #{filename}"
	reportfile.puts "<li>" + method_to_execute.call(filename) + "</li>"
    }
    reportfile.puts "</ul>"
    
    #space it out a bit
    puts ""
    reportfile.puts ""

    #start processing the other directories in current dir
    Dir.glob(current_directory.path + "/**") { |dirname|
	if File.directory? dirname
	    for_each_vb_file_in_dir(dirname, reportfile, method_to_execute)
	end
    }
end


#takes care of each detail page
def vb_file_report(filename)
    return if filename.nil? || filename.strip.empty?
    changed_filename = filename + ".htm"
    changed_filename[0] = $signature_survey_directory

    signature = ""
    lines = 0
    methods = 0
    indent_level = 0
    indent = lambda {"  " * indent_level}
	    
    File.open(filename, 'r') { |f|
        File.open(changed_filename, 'w+') { |g|
	    
	    #extract only relevant text
	    detail_page_text = ""
	    continued_function_sig = false;
            f.each_line { |line|
        	lines += 1

		# begin VB-specific code
                # this is pretty hacky for now, lots of edge cases that aren't
                # covered. Doing this really properly probably requires
                # something a bit more sophisticated
		if continued_function_sig
        	    detail_page_text.concat indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[] 
		    continued_function_sig = false unless line =~ /_$/ 
		elsif line =~ /^\s+'/
		    signature += "'"
		    detail_page_text.concat "'"
        	elsif line =~ /#Region/i
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[]
        	    indent_level += 1
		    signature += "."
        	elsif line =~ /#End Region/i
        	    indent_level -= 1
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[]
		    signature += "."
		elsif line =~ /(End\s+Sub|End\s+Function)/i
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[] 
		    signature += "}"
		elsif line =~ /Exit Sub/i
		    #ugly special case; treat like normal line
        	    detail_page_text.concat "."
		    signature += "."
        	elsif line =~ /(\bSub\s|\bFunction\s)/i
        	    methods += 1
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[] 
		    signature += "{"
		    continued_function_sig = true if line =~ /_$/ 
		elsif line =~ /Next|End\s+While|^\s+Loop/i
		    signature += "]"
		    detail_page_text.concat "]"
		elsif line =~ /^\s+For Each|^\s+While|^\s+For\s|^\s+Do Until