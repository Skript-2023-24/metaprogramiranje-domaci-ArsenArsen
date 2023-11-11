# frozen_string_literal: true

# A terrible toy made up of pure crystalline Ruby
# Copyright (C) 2023  Arsen Arsenović <aarsenovic8422rn@raf.rs>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

require "forwardable"
require "pp"

# A column-based table
class Worksheet
  include Enumerable
  attr_reader :table

  def initialize(matrix)
    @table = matrix
    return if @table.is_a? Hash

    @table = @table
             .filter { |row| row.none? { |i| /(sub)?total/i.match? i } }
             .filter { |row| row.any? { |i| !i.strip.empty? } }
             .transpose
             .filter { |col| col.any? { |i| !i.strip.empty? } }
             .to_h { |i| f, *r = i; [f, r] }
  end

  def self.from_worksheet(wsheet)
    Worksheet.new(wsheet.rows)
  end

  def row(idx)
    return table.keys if idx.zero?

    idx -= 1 # Zero-indexed contents.
    table.map { |_j, i| i[idx] }
  end

  def rows
    table
      .to_a
      .map { |x| h, t = x; [h, *t] }
      .transpose
  end

  def each(&)
    rows.flatten(1).each(&)
  end

  def -(other)
    raise "Can only subtract another sheet" unless other.is_a? Worksheet
    raise "Sheet headers differ" unless other.table.keys == table.keys

    header = [row(0)]
    data = rows - other.rows
    Worksheet.new(header + data)
  end

  def +(other)
    raise "Can only subtract another sheet" unless other.is_a? Worksheet
    raise "Sheet headers differ" unless other.table.keys == table.keys

    data = rows + other.rows.drop(1)
    Worksheet.new data
  end

  # TODO(arsen): merged fields?

  def [](name)
    col = table[name]
    Column.new(self, col) unless col.nil?
  end

  def respond_to_missing?(_, *)
    true
  end

  def method_missing(name, *args)
    raise "Function #{name} does not accept any arguments" unless args.empty?

    self.[] name.to_s
  end
end

def integery?(str)
  Integer(str)
  true
rescue StandardError
  false
end

# A single column of a worksheet.
class Column
  include Enumerable
  extend Forwardable

  def initialize(sheet, column)
    raise "expected non-nil column and sheet" if column.nil? || sheet.nil?

    @sheet = sheet
    @column = column
  end

  def each(&)
    @column.each(&)
  end

  def sum
    map(&:to_i).reduce(&:+)
  end

  def avg
    sum.to_f / (@column.count { |i| integery? i })
  end

  def_delegators :@column, :[], :[]=

  def respond_to_missing?(_, *)
    true
  end

  def method_missing(name, *args)
    raise "Function #{name} does not accept any arguments" unless args.empty?

    name = name.to_s
    # One-indexed
    @sheet.row 1 + @column.index { |i| name == i }
  end
end

require "google_drive"
session = GoogleDrive::Session.from_config "config.json"

gsheet = session.spreadsheet_by_url "https://docs.google.com/spreadsheets/d/1NbTJHnaEKHzo0CByipPJ0DhSzAiZY0MeePOE4RmVYOA/edit"

wsheet = gsheet.worksheets.find { 1 }
p wsheet.rows
ws = Worksheet.from_worksheet wsheet

def section
  yield
  puts "--------------------------- 8< ---------------------------"
end

section do
  pp ws
end

section do
  pp ws.row 1
end

section do
  ws.each { |x| pp x }
end

section do
  id = ws["ID"]
  pp id
  pp id[1]
  pp id[1] = "20"
  pp id[1]
end

section do
  pp ws.ID
  pp (ws.ID.filter { |i| i.to_i > 10 })
end

section do
  pp ws.ID.RN9
end

section do
  c = ws["Prva Kolona"]
  pp c.avg
  pp c.sum
end

section do
  ws2 = Worksheet.new ws.table.clone
  ws2.table.transform_values! { |i| i[1..2] }
  ws3 = pp ws - ws2
  puts
  pp ws - ws3
  puts
  pp ws3 + ws2
end
