#!/usr/bin/env ruby
# encoding: utf-8

begin
  require 'mathtype_to_mathml_plus'
rescue LoadError
  # Try to load from user install path if not found in system
  
  # For debug: print failing path
  # STDERR.puts "Ruby LoadPath: #{$LOAD_PATH}"
  raise
end

require 'base64'
require 'tempfile'

# Read binary data from STDIN (base64 encoded to avoid shell issues)
encoded_data = STDIN.read
if encoded_data.nil? || encoded_data.empty?
#   STDERR.puts "Error: No data received."
  exit 1
end

begin
  binary_data = Base64.decode64(encoded_data)
  
  # Workaround for ruby-ole/mathtype bug:
  # The mathtype gem calls Ole::Storage.open(arg, "rb+")
  # If arg is an IO object, ruby-ole raises ArgumentError because mode string is not allowed with IO.
  # So we must pass a file path.
  
  Tempfile.create(['mathtype', '.bin']) do |temp_file|
    temp_file.binmode
    temp_file.write(binary_data)
    temp_file.flush
    # We don't close immediately or it might be deleted depending on system/impl, 
    # but create with block usually yields an open file. 
    # Ole::Storage.open usually wants a path or an IO. 
    # If we pass path, it opens it.
    
    # Passing path to avoid the IO mode issue
    mathml = MathTypeToMathMLPlus::Converter.new(temp_file.path).convert
    puts mathml
  end
rescue => e
  # STDERR.puts "Ruby Error: #{e.message}"
  # STDERR.puts e.backtrace
  
  # To propagate detailed error to Python
  raise e
end
