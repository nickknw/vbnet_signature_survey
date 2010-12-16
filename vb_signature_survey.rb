$signature_survey_directory = "signature_survey" 
$root_directory_name = File.basename(Dir.pwd)
$file_extension = "vb"

$html_style_block = %Q| 
	<style type="text/css">
	    body { font-family: georgia, serif; }
	    h1 { color: #111111 }
	    li { font-family: Consolas, monospace; font-size: 14px; margin-bottom:20px; list-style-type: none; word-wrap:break-word; }
	    .line_count { color: #000080; }
	    .method_count { color: #008080; }
	    .signature { font-size: 12px; }
	    .inside_method { background-color: #eeeeee; }
	    .detail_page { font-family: Consolas, monospace; font-size: 14px; word-wrap: break-word}
	    .method_name { background-color: #ccccee ; font-weight: bold; }
	    .type { background-color: 88ffaa; font-weight: bold; }
	</style>|

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
            f.each_line { |line|
        	lines += 1

		#begin VB-specific code
        	if line =~ /#Region/
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[]
        	    indent_level += 1
		    signature += "."
        	elsif line =~ /#End Region/ 
        	    indent_level -= 1
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[]
		    signature += "."
		elsif line =~ /(End\s+Sub|End\s+Function)/
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[] 
		    signature += "}"
        	elsif line =~ /(Sub\s|Function\s)/
        	    methods += 1
        	    detail_page_text.concat "<br>" + indent[] + line.strip.gsub(/</, "&lt;").gsub(/>/, "&gt;") + "<br>" + indent[] 
		    signature += "{"
		elsif line =~ /^\s*$/
		    #ignore blank lines
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

	    signature_html = signature.gsub(/{/, "<span class='inside_method'>{")
	    signature_html = signature_html.gsub(/}/, "}</span>")
	    signature_html = "<span class='signature'>#{signature_html}</span>"

	    file_summary_line = methods_html + lines_html + file_name_html + signature_html

	    #write text to html file
	    g.puts %Q|<html><head><title>#{just_name_of_file}#{$html_style_block}</head><body><div class='detail_page'>|
	    g.puts "<h2>" + just_name_of_file + ": " + methods_html + lines_html + "</h2>"

	    detail_page_text.gsub!(/\s/, "&nbsp;")
	    detail_page_text.gsub!(/\b\w+\s?(?=\()/, "<span class='method_name'>\\0</span>")
	    detail_page_text.gsub!(/As&nbsp;(\w+\.)*\w+/, "<span class='type'>\\0</span>")

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
    report_file.puts %Q|<html><head>#{$html_style_block}</head><body>|
    report_file.puts "<h2>Using: #{File.expand_path(".")}</h2>"
    
    for_each_vb_file_in_dir(".", report_file, method(:vb_file_report))

    report_file.puts %Q|</body></html>|
}

puts "Report file saved as: ./#{$signature_survey_directory}/#{report_file_name_}"
