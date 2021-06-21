# Menu Title: Get Fonts for PDF's
# Needs Case: true
# Needs Selected Items: false
# Author: Cameron Stiller
# Version: 1.0
# Comment: Please note this has been provided as a guide only and should be thoroughly tested before entering production use.


nodeClasses={
	"com.aspose.pdf.Document"=>"A document object that, as the root of the document tree, provides access to the entire pdf document.",
	"com.aspose.pdf.Page"=>"A page of pdf document.",
	"com.aspose.pdf.Font"=>"Generic Font",
	"javax.swing.JFileChooser"=> "GUI element to open a file/folder chooser dialog",
	"javax.swing.JDialog"=> "Generic Dialog Window",
	"javax.swing.JPanel"=> "Generic Panel on a Window",
	"javax.swing.JFrame"=> "GUI grouping on a panel",
	"javax.swing.JProgressBar"=> "GUI element to display progress",
	"javax.swing.JOptionPane"=>"simple GUI pop"
}

require 'java'
require 'fileutils'
nodeClasses.each do | classname, description|
	begin
		#puts "Loading class: #{classname}\n\t#{description}"
		java_import classname
	rescue Exception => ex
		puts ex.message
		puts ex.backtrace
	end
end

ba=$utilities.getBulkAnnotater()

def show_message(message,title="Message")
    JOptionPane.showMessageDialog(nil,message,title,JOptionPane::PLAIN_MESSAGE)
end

def choose_dir (title="Choose Directory",loaddir="")
	dir = nil
	chooser = javax.swing.JFileChooser.new
	chooser.dialog_title = title
	chooser.file_selection_mode = JFileChooser::DIRECTORIES_ONLY
	chooser.setCurrentDirectory(java.io.File.new("#{loaddir}"))
	if chooser.show_open_dialog(nil) == JFileChooser::APPROVE_OPTION
		dir = chooser.selected_file.path.gsub('\\', '/')
	end
	return dir
end

class ProgressDialog < JDialog
	@spin_thread=Thread.new(){sleep(1)}
	def initialize(value,min,max,title="Progress Dialog",width=800,height=80)
		@spin_mutex = Mutex.new
		@progress_stat = JProgressBar.new  min, max
		self.setValue(value)
		@exportthread=nil
		super nil, true
		body=JPanel.new(java.awt.GridLayout.new(0,1))
		body.add(@progress_stat)
		self.add(body)
		self.setDefaultCloseOperation JFrame::DISPOSE_ON_CLOSE
		self.setTitle(title)
		self.setDefaultCloseOperation JFrame::DISPOSE_ON_CLOSE
		self.setSize width, height
		self.setLocationRelativeTo nil
		Thread.new{
				yield self
			sleep(0.2)
			self.dispose()
		}
		self.setVisible true
	end

	def setValue(value=0)
		spin(false)
		@progress_stat.setValue(value)
	end

	def spin(start) #start true will start spinning, start false will stop spinning
		@spin_mutex.synchronize{
			if(@spin_thread.nil?)
				@spin_thread=Thread.new(){}
				sleep(0.1)
			end
			if(@spin_thread.alive? ==true)
				if(start)
					@spin_thread.kill()
					@spin_thread=Thread.new(){endless()}
				else
					@spin_thread.kill()
					@spin_thread=Thread.new(){}
				end
			else
				if(start)
					@spin_thread=Thread.new(){ endless()}
				end
			end
		}
	end

	def endless()
		begin
		loop{
			0.upto(@progress_stat.getMaximum()) do |i|
				@progress_stat.setValue(i)
				sleep(0.1)
			end
		}
		rescue Exception => ex
			puts ex.message
		end
	end

	def setMax(max)
		spin(false)
		@progress_stat.setMaximum(max)
	end
end



title="Temp Directory"
initialpath= $current_case.getBatchLoads().last().getParallelProcessingSettings()["Worker temp directory"]
$tempDir=choose_dir(title,initialpath)
if($tempDir.nil?)
	show_message("no Temp selected")
	exit
end
if($tempDir=="")
	show_message("no Temp selected")
	exit
end

itemcount=0
tag_font_cloud=Hash.new()
additional_messaging=""
ProgressDialog.new(0,0,100,"initialising",400,80) do | dialog|
	begin
		be=$utilities.getBinaryExporter()
		dialog.setTitle("Searching for all PDF items")
		dialog.spin(true)
		pdf_items=[]
		if($current_selected_items.length > 0)
			pdf_items=$current_selected_items.select{|item|item.getType().getName()=="application/pdf"}
			additional_messaging=" selected "
		else
			pdf_items=$current_case.searchUnsorted("mime-type:\"application/pdf\" NOT (properties:FailureDetail AND NOT flag:encrypted)")
		end
		dialog.setTitle("Extracting fonts from #{pdf_items.length} pdf items")
		dialog.setMax(pdf_items.length)
		itemcount=pdf_items.length
		fonts=Array.new()
		pdf_items.each_with_index do | item,index |
			dialog.setValue(index)
			dump_location=$tempDir + "\\" + item.getGuid() + ".pdf"
			be.exportItem(item, dump_location)
			pdfFile = Document.new(dump_location)
			begin
				pdfFile.getPages().to_a.each_with_index do | page,index|
					begin
						fonts.push *page.getResources().getFonts()
					rescue Exception => ex
						puts "Error getting fonts on page #{index} of #{item.getGuid()}"
						puts ex.message
						puts ex.backtrace
					end
				end
				fonts.map(&:getFontName).uniq.each do | font|
					if(!tag_font_cloud.has_key? font)
						tag_font_cloud[font]=[]
					end
					tag_font_cloud[font].push item
				end
			rescue Exception => ex
				puts "Error iterating pages of #{item.getGuid()}"
				puts ex.message
				puts ex.backtrace
			end
			pdfFile.close()
			FileUtils.rm(dump_location)
			
		end
	rescue Exception => ex
		puts ex.message
		puts ex.backtrace
	end
	dialog.setTitle("Tagging fonts")
	dialog.setMax(tag_font_cloud.keys.length)
	tag_font_cloud.keys.each_with_index do | tag,index|
		ba.addTag("Font\|#{tag.strip()}", tag_font_cloud[tag])
		dialog.setValue(index)
	end
end
show_message("Generated font properties for:\n\t#{itemcount} pdf #{additional_messaging} items\n\tFound Fonts:#{tag_font_cloud.keys.length}")
