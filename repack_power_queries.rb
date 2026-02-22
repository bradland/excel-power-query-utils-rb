#!/usr/bin/env ruby

require 'zip'
require 'base64'
require 'rexml/document'
require 'stringio'
require 'optparse'
require 'fileutils'

# = PowerQueryRepacker
#
# This class provides the inverse functionality of the +PowerQueryExtractor+.
# It takes a directory of unpacked Data Mashup files and injects them back 
# into a target Microsoft Excel (.xlsx) file.
#
# == The Repacking Process:
# 1.  **Inner Compression**: Re-archives the source directory (containing 
#     +Formulas/Section1.m+, etc.) into a ZIP binary.
# 2.  **Header Restoration**: Prepends the mandatory Microsoft-proprietary 
#     MS-QDEFF binary header. This header is dynamically harvested from the 
#     target template to ensure version compatibility.
# 3.  **OOXML Injection**: Encodes the final binary as Base64 and replaces 
#     the content of the +<DataMashup>+ element within the workbook's 
#     +customXml+ parts.
#
# == Example Usage (Library):
#   repacker = PowerQueryRepacker.new("./unpacked_dir", "original.xlsx", "updated.xlsx")
#   repacker.repack
#
class PowerQueryRepacker
  # The "Magic Number" signature for the start of a ZIP file archive.
  ZIP_MAGIC_NUMBER = "PK\x03\x04".freeze

  # Creates a new repacker instance.
  #
  # [source_dir]   The directory containing the unpacked DataMashup files 
  #                (must contain the +Formulas+ and +[Config]+ subdirectories).
  # [target_xlsx]  The original Excel file to use as a structural template.
  # [output_xlsx]  The path where the new, modified Excel file will be saved.
  def initialize(source_dir, target_xlsx, output_xlsx)
    @source_dir = File.expand_path(source_dir)
    @target_xlsx = target_xlsx
    @output_xlsx = output_xlsx
  end

  # Orchestrates the repacking and injection process.
  #
  # Returns +true+ if the new file was created successfully, +false+ otherwise.
  def repack
    unless Dir.exist?(@source_dir)
      warn "Error: Source directory '#{@source_dir}' not found."
      return false
    end

    unless File.exist?(@target_xlsx)
      warn "Error: Target template '#{@target_xlsx}' not found."
      return false
    end

    # 1. Rebuild the inner ZIP from the source directory
    inner_zip_buffer = create_inner_zip

    # 2. Extract original MS-QDEFF header to maintain binary integrity
    header = extract_original_header
    
    # 3. Combine header + ZIP and encode
    full_mashup_binary = header + inner_zip_buffer
    encoded_mashup = Base64.strict_encode64(full_mashup_binary)

    # 4. Inject into the OOXML structure
    update_xlsx_with_mashup(encoded_mashup)
    true
  rescue StandardError => e
    warn "An error occurred during repacking: #{e.message}"
    false
  end

  private

  # Walks the source directory and creates a ZIP binary stream.
  # Ensures internal paths (e.g. 'Formulas/Section1.m') are preserved.
  def create_inner_zip
    Zip::OutputStream.write_buffer do |zos|
      Dir.glob(File.join(@source_dir, "**", "*")).each do |file|
        next if File.directory?(file)
        
        # Calculate the relative path for the ZIP entry name
        entry_name = file.sub(%r{^#{@source_dir}/}, "")
        
        zos.put_next_entry(entry_name)
        zos.write(File.read(file))
      end
    end.string
  end

  # Scans the target Excel file to retrieve the existing MS-QDEFF header.
  # This header precedes the ZIP content in the DataMashup blob.
  def extract_original_header
    Zip::File.open(@target_xlsx) do |zip|
      zip.glob('customXml/item*.xml').each do |entry|
        doc = REXML::Document.new(entry.get_input_stream.read)
        element = REXML::XPath.first(doc, "//*[local-name()='DataMashup']")
        next unless element
        
        raw_bin = Base64.decode64(element.text)
        zip_idx = raw_bin.index(ZIP_MAGIC_NUMBER)
        return zip_idx ? raw_bin[0...zip_idx] : ""
      end
    end
    ""
  end

  # Copies the template to the output path and replaces the DataMashup content.
  def update_xlsx_with_mashup(new_base64)
    FileUtils.cp(@target_xlsx, @output_xlsx)
    
    # Use a temporary file to handle the ZIP update safely
    Zip::File.open(@output_xlsx) do |zip|
      zip.glob('customXml/item*.xml').each do |entry|
        content = entry.get_input_stream.read
        doc = REXML::Document.new(content)
        element = REXML::XPath.first(doc, "//*[local-name()='DataMashup']")
        next unless element

        element.text = new_base64
        
        # Replace the XML entry in the ZIP with updated text
        zip.get_output_stream(entry.name) { |f| f.write(doc.to_s) }
        puts "Injected updated DataMashup into #{entry.name}"
      end
    end
  end
end

# == Command Line Interface
#
# Usage:
#   ./pq_repack.rb [options] -s <unpacked_dir> -t <template.xlsx> -o <output.xlsx>
#
# Options:
#   -s, --source DIR    Directory containing the unpacked DataMashup structure.
#   -t, --target FILE   The original Excel file to use as a structural base.
#   -o, --output FILE   The path for the newly generated Excel file.
#   -h, --help          Displays the help menu.
if __FILE__ == $0
  options = {}

  parser = OptionParser.new do |opts|
    opts.banner = "Usage: #{File.basename($0)} [options]"
    opts.separator ""
    opts.separator "Options:"

    opts.on("-s", "--source DIR", "Directory with unpacked DataMashup files") do |v|
      options[:source] = v
    end

    opts.on("-t", "--target FILE", "Original Excel file (template)") do |v|
      options[:target] = v
    end

    opts.on("-o", "--output FILE", "Path for the new Excel file") do |v|
      options[:output] = v
    end

    opts.on_tail("-h", "--help", "Show this message") do
      puts opts
      exit
    end
  end

  begin
    parser.parse!
  rescue OptionParser::InvalidOption, OptionParser::MissingArgument => e
    warn e.message
    puts parser
    exit 1
  end

  if options[:source] && options[:target] && options[:output]
    repacker = PowerQueryRepacker.new(options[:source], options[:target], options[:output])
    success = repacker.repack
    exit(success ? 0 : 1)
  else
    puts parser
    exit 1
  end
end
