#!/usr/bin/env ruby

require 'win32ole'
require 'optparse'

# = PowerQueryRefresher
#
# This class provides automation for Microsoft Excel using the Win32OLE bridge.
# It is specifically designed to be used after repacking a Power Query 
# DataMashup into an .xlsx file.
#
# == The Refresh Process:
# 1.  **Automation Initialization**: Starts a headless instance of the 
#     Excel Application.
# 2.  **Synchronous Configuration**: Iterates through all Workbook 
#     connections to disable +BackgroundQuery+. This is critical because 
#     it forces Ruby to wait for the Power Query engine to finish 
#     fetching data before moving to the next line of code.
# 3.  **Refresh & Persistence**: Executes +RefreshAll+, saves the 
#     calculated results back to the file, and ensures the Excel 
#     process is terminated cleanly.
#
# == Requirements:
# *   **Operating System**: Windows (required for Win32OLE and Excel).
# *   **Software**: Microsoft Excel must be installed on the machine.
#
# == Example Usage (Library):
#   refresher = PowerQueryRefresher.new("updated_report.xlsx")
#   refresher.refresh
#
class PowerQueryRefresher
  # Creates a new refresher instance.
  #
  # [file_path] The path to the Excel workbook to be refreshed.
  def initialize(file_path)
    # Excel's COM object requires absolute paths with Windows-style backslashes
    @file_path = File.expand_path(file_path).gsub("/", "\\")
  end

  # Connects to Excel via OLE, refreshes all queries, and saves the file.
  #
  # Returns +true+ if the refresh and save were successful, +false+ otherwise.
  def refresh
    unless File.exist?(@file_path)
      warn "Error: File '#{@file_path}' not found."
      return false
    end

    begin
      # 1. Initialize Excel Application
      @excel = WIN32OLE.new('Excel.Application')
      @excel.Visible = false
      @excel.DisplayAlerts = false

      # 2. Open the Workbook
      puts "Opening #{@file_path}..."
      @workbook = @excel.Workbooks.Open(@file_path)

      # 3. Force Synchronous Refresh
      # We must disable BackgroundQuery so RefreshAll blocks the script execution
      # until the data is fully loaded.
      configure_synchronous_queries

      # 4. Perform Refresh
      puts "Refreshing all Power Query connections... (this may take a while)"
      @workbook.RefreshAll

      # 5. Save and Close
      @workbook.Save
      puts "Refresh successful. Workbook saved."
      true
    rescue StandardError => e
      warn "An error occurred during Excel automation: #{e.message}"
      false
    ensure
      cleanup
    end
  end

  private

  # Iterates through OLEDB and Workbook connections to disable background processing.
  def configure_synchronous_queries
    @workbook.Connections.each do |conn|
      # Type 1 = OLEDB (Power Query), Type 2 = ODBC
      if conn.Type == 1 || conn.Type == 2
        begin
          conn.OLEDBConnection.BackgroundQuery = false
        rescue
          # Some connections might not support this property; skip if error
          next
        end
      end
    end
  end

  # Ensures the Excel process is terminated and memory is released.
  def cleanup
    if @workbook
      @workbook.Close(false)
      @workbook = nil
    end
    if @excel
      @excel.Quit
      @excel = nil
    end
    # Explicitly trigger GC to help release COM objects
    GC.start
  end
end

# == Command Line Interface
#
# Usage:
#   ./pq_refresh.rb <excel_file.xlsx>
#
# Options:
#   -h, --help    Show the help menu.
if __FILE__ == $0
  options = {}
  
  parser = OptionParser.new do |opts|
    opts.banner = "Usage: #{File.basename($0)} <excel_file.xlsx>"
    opts.separator ""
    opts.separator "Note: This script requires Windows and Microsoft Excel."

    opts.on_tail("-h", "--help", "Show this message") do
      puts opts
      exit
    end
  end

  begin
    parser.parse!
  rescue OptionParser::InvalidOption => e
    warn e.message
    puts parser
    exit 1
  end

  if ARGV.empty?
    puts parser
    exit 1
  end

  # Instantiate and execute
  refresher = PowerQueryRefresher.new(ARGV.first)
  success = refresher.refresh
  exit(success ? 0 : 1)
end
