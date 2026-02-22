#!/usr/bin/env ruby

require 'zip'
require 'base64'
require 'rexml/document'
require 'stringio'
require 'optparse'
require 'fileutils'

# = PowerQueryExtractor
#
# This class provides functionality to extract the internal "Data Mashup" 
# content from Microsoft Excel OOXML (.xlsx) files. 
#
# Power Query logic is stored as a Base64-encoded binary stream within 
# +customXml/item[N].xml+. This binary contains a Microsoft-proprietary 
# header (MS-QDEFF) followed by a ZIP archive.
#
# == Example Usage (Library):
#   extractor = PowerQueryExtractor.new("data.xlsx", "out_dir", split_queries: true)
#   extractor.extract_all
#
class PowerQueryExtractor
  # The "Magic Number" signature for the start of a ZIP file archive.
  ZIP_MAGIC_NUMBER = "PK\x03\x04".freeze

  # Creates a new extractor instance.
  #
  # [file_path]      The path to the source .xlsx file.
  # [output_dir]     The directory where extracted files will be written.
  # [split_queries]  Boolean; if true, parses Section1.m into individual .m files.
  def initialize(file_path, output_dir = ".", split_queries: false)
    @file_path = Array(file_path).first # Handle potential ARGV array
    @output_dir = output_dir
    @split_queries = split_queries
  end

  # Orchestrates the extraction process.
  #
  # Opens the Excel ZIP, locates the DataMashup XML element, decodes the 
  # Base64 string, strips the MS-QDEFF header, and unpacks the internal 
  # ZIP contents. If +@split_queries+ is true, it further processes +Section1.m+.
  #
  # Returns +true+ if successful, +false+ otherwise.
  def extract_all
    unless File.exist?(@file_path)
      warn "Error: File '#{@file_path}' not found."
      return false
    end

    FileUtils.mkdir_p(@output_dir) unless Dir.exist?(@output_dir)

    Zip::File.open(@file_path) do |excel_zip|
      excel_zip.glob('customXml/item*.xml').each do |entry|
        xml_content = entry.get_input_stream.read
        doc = REXML::Document.new(xml_content)
        
        mashup_element = REXML::XPath.first(doc, "//*[local-name()='DataMashup']")
        next unless mashup_element

        full_binary = Base64.decode64(mashup_element.text)
        zip_start_index = full_binary.index(ZIP_MAGIC_NUMBER)
        
        if zip_start_index
          zip_binary = full_binary[zip_start_index..-1]
          unpack_mashup(zip_binary)
          split_m_logic if @split_queries
          return true
        end
      end
    end
    warn "No DataMashup contents found in '#{@file_path}'."
    false
  rescue StandardError => e
    warn "An error occurred during extraction: #{e.message}"
    false
  end

  private

  # Unpacks the inner ZIP binary found within the DataMashup stream.
  # Maintains the internal directory structure during extraction.
  def unpack_mashup(binary_data)
    Zip::InputStream.open(StringIO.new(binary_data)) do |io|
      while (entry = io.get_next_entry)
        dest_path = File.join(@output_dir, entry.name)
        FileUtils.mkdir_p(File.dirname(dest_path))
        
        File.open(dest_path, 'wb') { |f| f.write(io.read) }
        puts "Extracted: #{entry.name} -> #{dest_path}"
      end
    end
  end

  # Parses +Formulas/Section1.m+ and extracts each 'shared' query into 
  # a separate file within an +Individual_Queries+ subdirectory.
  def split_m_logic
    m_file = File.join(@output_dir, 'Formulas', 'Section1.m')
    unless File.exist?(m_file)
      warn "Warning: Section1.m not found, skipping split."
      return
    end

    content = File.read(m_file)
    queries_dir = File.join(@output_dir, 'Individual_Queries')
    FileUtils.mkdir_p(queries_dir)

    # Regex captures 'shared QueryName = ...;' patterns
    # Handles both standard names and #"Quoted Name" identifiers
    query_regex = /shared\s+(?:#"(?<quoted>[^"]+)"|(?<unquoted>\w+))\s*=\s*(?<logic>.*?);\s*(?=shared|\z)/m

    content.scan(query_regex).each do |quoted, unquoted, logic|
      raw_name = (quoted || unquoted)
      safe_name = raw_name.gsub(/[^0-9A-Za-z.\- ]/, '_')
      
      File.write(File.join(queries_dir, "#{safe_name}.m"), logic.strip)
      puts "Split: #{raw_name} -> #{queries_dir}/#{safe_name}.m"
    end
  end
end

# == Command Line Interface
#
# Usage:
#   ./pq_extractor.rb [options] <excel_file.xlsx>
#
# Options:
#   -o, --output DIRECTORY   Specify output directory (default: current).
#   -s, --split              Extract individual queries from Section1.m.
#   -h, --help               Show the help menu.
if __FILE__ == $0
  options = { output: ".", split: false }

  parser = OptionParser.new do |opts|
    opts.banner = "Usage: #{File.basename($0)} [options] <excel_file.xlsx>"
    opts.separator ""
    opts.separator "Options:"

    opts.on("-o", "--output DIRECTORY", "Directory to unpack files into (default: current)") do |dir|
      options[:output] = dir
    end

    opts.on("-s", "--split", "Parse Section1.m into individual .m query files") do
      options[:split] = true
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

  if ARGV.empty?
    puts parser
    exit 1
  end

  extractor = PowerQueryExtractor.new(ARGV, options[:output], split_queries: options[:split])
  success = extractor.extract_all
  exit(success ? 0 : 1)
end
