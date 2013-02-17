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
		elsif line =~ /^\s+For Each|^\s+While|^\s+For\s|^\s+Do Until/i
		    signature += "["
		    detail_page_text.concat "["
		elsif line =~ /End\s+If|End\s+Select/i
		    signature += ")"
		    detail_page_text.concat ")"
                elsif line =~ /\s+Else/i
                    signature += "|"
                    detail_page_text.concat "|"
		elsif line =~ /If.*Then\s*$|Select Case/i
		    signature += "("
		    detail_page_text.concat "("
                elsif line =~ /\s+Case\s+/i
                    signature += "|"
                    detail_page_text.concat "|"
		elsif line =~ /If/i
		    signature += "?"
		    detail_page_text.concat "?"
		elsif line =~ /^\s*$/
        	    detail_page_text.concat " "
		    signature += " "
        	else
        	    detail_page_text.concat "."
		    signature += "."
        	end
		#end VB-specific code
            }

	    #format return value info into html
	    just_name_of_file = File.basename(changed_filename).sub(/\..+/, '')
	    spaces_in_justified_name = just_name_of_file.ljust(20).gsub(/\S/, '').gsub(' ', "&nbsp;")

	    methods_html = "<span class='method_count'>" + "#{methods}m".rjust(3).gsub(' ', "&nbsp;") + "</span>"
	    lines_html = " <span class='line_count'>" + " #{lines}L".ljust(8).gsub(' ', "&nbsp;") + "</span>"
	    file_name_html = "<a href=\"#{filename + ".htm"}\">#{just_name_of_file}</a>#{spaces_in_justified_name} "

	    signature_html = signature
	    signature_html.gsub!(/'/, "<span class='comment'>'</span>")
	    signature_html.gsub!(/{/, "<div class='inside_method'>{")
	    signature_html.gsub!(/}/, "}</div>")

	    signature_html.gsub!(/\[/, "<p class='inside_loop'>[")
	    signature_html.gsub!(/\]/, "]</p>")

	    signature_html.gsub!(/\(/, "<span class='inside_if'>(")
	    signature_html.gsub!(/\)/, ")</span>")
	    signature_html.gsub!(/\?/, "<span class='inside_if'>?</span>")

	    signature_html = "<span class='signature'>#{signature_html}</span>"

	    file_summary_line = methods_html + lines_html + file_name_html + signature_html

	    #write text to html file
	    file_path_without_extension = filename.sub(/(\/.*?)\..+$/, "\\1")
	    g.puts %Q|<html><head><title>#{file_path_without_extension}</title>#{$html_style_block}</head><body>|
	    g.puts "<h2>" + file_path_without_extension + ": " + methods_html + lines_html + "</h2><div class='detail_page'>"
	    g.puts $legend_html

	    detail_page_text.gsub!(/'/, "<span class='comment'>'</span>")
	    detail_page_text.gsub!(/>\s*\(\) (As|Handles)/, ">&#40;&#41; \\1")
	    detail_page_text.gsub!(/\) (As|Handles)/, "&#41; \\1")
	    detail_page_text.gsub!(/\s\s/, "&nbsp;&nbsp;")

	    detail_page_text.gsub!(/(Function|Sub)\s+([^(]+)\(/, "\\1 <span class='method_name'>\\2</span>&#40;")
	    detail_page_text.gsub!(/>\s*\((ByVal|ByRef)/, ">&#40;\\1")
	    detail_page_text.gsub!(/As (\w+\.)*\w+/, "<span class='type'>\\0</span>")

	    detail_page_text.gsub!(/{/, "<div class='inside_method'>{")
	    detail_page_text.gsub!(/}/, "}</div>")

	    detail_page_text.gsub!(/\[/, "<p class='inside_loop'>[")
	    detail_page_text.gsub!(/\]/, "]</p>")

	    detail_page_text.gsub!(/\(/, "<span class='inside_if'>(")
	    detail_page_text.gsub!(/\)/, ")</span>")

	    detail_page_text.gsub!(/\?/, "<span class='inside_if'>?</span>")

	    g.puts detail_page_text

	    g.puts "</div></body></html>"

	    return file_summary_line
        }
    }

end

#----start of execution----

#make sure the directory that the report should sit in exists
Dir.mkdir $signature_survey_directory unless Dir.exists? $signature_survey_directory

#set report file name
report_file_name_ = "#{$root_directory_name}-report-vb.htm"

#start generating the report
File.open($signature_survey_directory + "/" + report_file_name_, "w+") { |report_file|
    puts "Using: #{File.expand_path(".")}"
    puts ""
    report_file.puts %Q|<html><head><title>VB Signature Survey</title>#{$html_style_block}</head><body>|
    report_file.puts "<h1>VB Signature Survey</h1>"
    report_file.puts "<h2>Using: #{File.expand_path(".")}</h2>"
    report_file.puts $legend_html
    
    for_each_vb_file_in_dir(".", report_file, method(:vb_file_report))

    report_file.puts %Q|</body></html>|
}

puts "Report file saved as: ./#{$signature_survey_directory}/#{report_file_name_}"
