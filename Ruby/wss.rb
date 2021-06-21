$temp=ENV_JAVA["java.io.tmpdir"]
if($temp.to_s=="")
	$temp="C:/Temp"
end

def nuix_worker_item_callback_init	
	nodeClasses={
		"com.aspose.pdf.Document"=>"A document object that, as the root of the document tree, provides access to the entire pdf document.",
		"com.aspose.pdf.Page"=>"A page of pdf document.",
		"com.aspose.pdf.Font"=>"Generic Font",
	}

	require 'java'
	require 'fileutils'
	nodeClasses.each do | classname, description|
		begin
			puts "Loading class: #{classname}\n\t#{description}"
			java_import classname
		rescue Exception => ex
			puts ex.message
			puts ex.backtrace
		end
	end
end

def nuix_worker_item_callback(worker_item)
	source_item=worker_item.getSourceItem()
	if(source_item.getType.getName=="application/pdf")
		status=false
		item_guid=worker_item.getGuidPath().last()
		path=$temp + "/" + item_guid
		begin
			source_item.getBinary().getBinaryData().copyTo(path)
		rescue Exception => ex
			puts "Error extracting binary to temp location #{path}"
			puts ex.message
			puts ex.backtrace
		end
		fonts=Array.new()
		
		begin
			pdfFile = Document.new(path)
			begin
				pdfFile.getPages().to_a.each_with_index do | page,index|
					begin
						fonts.push *page.getResources().getFonts()
					rescue Exception => ex
						puts "Error getting fonts on page #{index} of #{item_guid}"
						puts ex.message
						puts ex.backtrace
					end
				end
				fonts.map(&:getFontName).uniq.each do | font|
					worker_item.addTag("Font\|#{font.strip()}")
				end
			rescue Exception => ex
				puts "Error iterating pages of #{item_guid}"
				puts ex.message
				puts ex.backtrace
			end
			pdfFile.close()
			FileUtils.rm(path)
			status=true
		rescue Exception => ex
			puts "Errors reading pdf file #{item_guid}"
			puts ex.message
			puts ex.backtrace
		end
		if(!(status))
			worker_item.addTag("Font ERROR (see logs)")
		end
	end
end