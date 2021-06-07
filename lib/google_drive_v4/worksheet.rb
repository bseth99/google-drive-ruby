# Author: Hiroshi Ichikawa <http://gimite.net/>
# The license of this source is "New BSD Licence"

require "set"

require "google_drive_v4/util"
require "google_drive_v4/error"
require "google_drive_v4/table"
require "google_drive_v4/list"


module GoogleDriveV4

    # A worksheet (i.e. a tab) in a spreadsheet.
    # Use GoogleDrive::Spreadsheet#worksheets to get GoogleDrive::Worksheet object.
    class Worksheet

        include(Util)

        def initialize(session, spreadsheet, cells_feed_url, title = nil, updated = nil) #:nodoc:

          @session = session
          @spreadsheet = spreadsheet
          @cells_feed_url = cells_feed_url
          @title = title
          @updated = updated

          @cells = nil
          @input_values = nil
          @numeric_values = nil
          @modified = Set.new()
          @list = nil
          @properties = Hash[]

        end

        # URL of cell-based feed of the worksheet.
        attr_reader(:cells_feed_url)
        attr_accessor( :properties )

        def divine_spreadsheet_id
           if !(@cells_feed_url =~ %r{^https?://sheets.googleapis.com/v4/spreadsheets/(.*)/values/(.*)$})
            raise(GoogleDriveV4::Error,
              "Cells feed URL is in unknown format: #{@cells_feed_url}")
           end
           return "#{$1}"
        end

        # URL of worksheet feed URL of the worksheet.
        def worksheet_feed_url
          # I don't know good way to get worksheet feed URL from cells feed URL.
          # Probably it would be cleaner to keep worksheet feed URL and get cells feed URL
          # from it.
          if !(@cells_feed_url =~ %r{^https?://sheets.googleapis.com/v4/spreadsheets/(.*)/values/(.*)$})
            raise(GoogleDriveV4::Error,
              "Cells feed URL is in unknown format: #{@cells_feed_url}")
          end
          return "https://sheets.googleapis.com/v4/spreadsheets/#{self.divine_spreadsheet_id}/"
        end

        # GoogleDrive::Spreadsheet which this worksheet belongs to.
        def spreadsheet
          if !@spreadsheet
            @spreadsheet = @session.spreadsheet_by_key( self.divine_spreadsheet_id )
          end
          return @spreadsheet
        end

        # Returns content of the cell as String. Arguments must be either
        # (row number, column number) or cell name. Top-left cell is [1, 1].
        #
        # e.g.
        #   worksheet[2, 1]  #=> "hoge"
        #   worksheet["A2"]  #=> "hoge"
        def [](*args)
          (row, col) = parse_cell_args(args)
          return self.cells[[row, col]] || ""
        end

        # Updates content of the cell.
        # Arguments in the bracket must be either (row number, column number) or cell name.
        # Note that update is not sent to the server until you call save().
        # Top-left cell is [1, 1].
        #
        # e.g.
        #   worksheet[2, 1] = "hoge"
        #   worksheet["A2"] = "hoge"
        #   worksheet[1, 3] = "=A1+B1"
        def []=(*args)
          (row, col) = parse_cell_args(args[0...-1])
          value = args[-1].to_s()
          reload() if !@cells
          @cells[[row, col]] = value
          @input_values[[row, col]] = value
          #@numeric_values[[row, col]] = nil
          @modified.add([row, col])
          self.max_rows = row if row > @max_rows
          self.max_cols = col if col > @max_cols
          if value.empty?
            @num_rows = nil
            @num_cols = nil
          else
            @num_rows = row if row > num_rows
            @num_cols = col if col > num_cols
          end
        end

        # Updates cells in a rectangle area by a two-dimensional Array.
        # +top_row+ and +left_col+ specifies the top-left corner of the area.
        #
        # e.g.
        #   worksheet.update_cells(2, 3, [["1", "2"], ["3", "4"]])
        def update_cells(top_row, left_col, darray)
          darray.each_with_index() do |array, y|
            array.each_with_index() do |value, x|
              self[top_row + y, left_col + x] = value
            end
          end
        end

        # Returns the value or the formula of the cell. Arguments must be either
        # (row number, column number) or cell name. Top-left cell is [1, 1].
        #
        # If user input "=A1+B1" to cell [1, 3]:
        #   worksheet[1, 3]              #=> "3" for example
        #   worksheet.input_value(1, 3)  #=> "=RC[-2]+RC[-1]"
        def input_value(*args)
           raise "No"
          (row, col) = parse_cell_args(args)
          reload() if !@cells
          return @input_values[[row, col]] || ""
        end

        # Returns the numeric value of the cell. Arguments must be either
        # (row number, column number) or cell name. Top-left cell is [1, 1].
        #
        # e.g.
        #   worksheet[1, 3]                #=> "3,0" # it depends on locale, currency...
        #   worksheet.numeric_value(1, 3)  #=> 3.0
        #
        # Returns nil if the cell is empty or contains non-number.
        #
        # If you modify the cell, its numeric_value is nil until you call save() and reload().
        #
        # For details, see:
        # https://developers.google.com/google-apps/spreadsheets/#working_with_cell-based_feeds
        def numeric_value(*args)
           raise "No"
          (row, col) = parse_cell_args(args)
          reload() if !@cells
          return @numeric_values[[row, col]]
        end

        # Row number of the bottom-most non-empty row.
        def num_rows
          reload() if !@cells
          # Memoizes it because this can be bottle-neck.
          # https://github.com/gimite/google-drive-ruby/pull/49
          return @num_rows ||= @input_values.select(){ |(r, c), v| !v.empty? }.map(){ |(r, c), v| r }.max || 0
        end

        # Column number of the right-most non-empty column.
        def num_cols
          reload() if !@cells
          # Memoizes it because this can be bottle-neck.
          # https://github.com/gimite/google-drive-ruby/pull/49
          return @num_cols ||= @input_values.select(){ |(r, c), v| !v.empty? }.map(){ |(r, c), v| c }.max || 0
        end

        # Number of rows including empty rows.
        def max_rows
          reload() if !@cells
          return @max_rows
        end

        # Updates number of rows.
        # Note that update is not sent to the server until you call save().
        def max_rows=(rows)
          reload() if !@cells
          @max_rows = rows
          @meta_modified = true
        end

        # Number of columns including empty columns.
        def max_cols
          reload() if !@cells
          return @max_cols
        end

        # Updates number of columns.
        # Note that update is not sent to the server until you call save().
        def max_cols=(cols)
          reload() if !@cells
          @max_cols = cols
          @meta_modified = true
        end

        # Title of the worksheet (shown as tab label in Web interface).
        def title
          reload() if !@title
          return @title
        end

         # Date updated of the worksheet (shown as tab label in Web interface).
        def updated
          reload() if !@updated
          return @updated
        end

        # Updates title of the worksheet.
        # Note that update is not sent to the server until you call save().
        def title=(title)
          reload() if !@cells
          @title = title
          @meta_modified = true
        end

        def cells #:nodoc:
          reload() if !@cells
          return @cells
        end

        # An array of spreadsheet rows. Each row contains an array of
        # columns. Note that resulting array is 0-origin so:
        #
        #   worksheet.rows[0][0] == worksheet[1, 1]
        def rows(skip = 0)
          nc = self.num_cols
          result = ((1 + skip)..self.num_rows).map() do |row|
            (1..nc).map(){ |col| self[row, col] }.freeze()
          end
          return result.freeze()
        end

        def get_sheet_data
           uri = URI(@cells_feed_url )
           uri.query = URI.encode_www_form( Hash[ "valueRenderOption" => "UNFORMATTED_VALUE", "dateTimeRenderOption" => "FORMATTED_STRING" ])
           @session.request(:get, uri.to_s , :header => Hash[ "Content-Type" => "application/json;charset=utf-8"], :response_type => :json )
        end

        # Reloads content of the worksheets from the server.
        # Note that changes you made by []= etc. is discarded if you haven't called save().
        def reload
           doc = self.get_sheet_data

           @max_rows = @properties["gridProperties"]["rowCount"]
           @max_cols = @properties["gridProperties"]["columnCount"]
           @title = @properties["title"]

           @num_cols = nil
           @num_rows = nil

           @cells = {}
           @input_values = {}
           if !doc["values"].nil?
              doc["values"].each_with_index do |entry, row|
                entry.each_with_index do |val, col|
                   if doc["majorDimension"] == "ROWS"
                      @cells[[row+1,col+1]] = val.to_s
                      @input_values[[row+1,col+1]] = val.to_s
                   elsif doc["majorDimension"] == "COLUMNS"
                      @cells[[col+1,row+1]] = val.to_s
                      @input_values[[col+1,row+1]] = val.to_s
                   else
                      raise(GoogleDriveV4::Error, "Unknown major dimension: %s" % doc["majorDimension"] )
                   end
                end
             end
          end

          @modified.clear()
          @meta_modified = false
          return true

        end

        def batch_update_url
           "https://sheets.googleapis.com/v4/spreadsheets/%s:batchUpdate" % self.divine_spreadsheet_id
        end

        def batch_values_url
           "https://sheets.googleapis.com/v4/spreadsheets/%s/values:batchUpdate" % self.divine_spreadsheet_id
        end

        def col_letter col
           acol = col - 1
           modcol = acol % 26
           divcol = acol / 26
           str = [ "A".ord + modcol ].pack("c")

           str = ( divcol > 0 ? [ "A".ord - 1 + divcol ].pack("c") : "" )  + str

           str
        end

        def cell_name row, col
           row_num = row.to_s
           col_let = self.col_letter col

           "'#{@title}'!#{col_let}#{row_num}"
        end

        def save
           sent = false

           if @meta_modified
             propupd = Hash[
               "requests" => [
                  Hash[
                     "updateSheetProperties" => Hash[
                        "properties" => Hash[
                           "sheetId" => @properties["sheetId"],
                           "title" => self.title,
                           "gridProperties" => Hash[
                              "rowCount" => self.max_rows,
                              "columnCount" => self.max_cols
                           ]
                        ],
                        "fields" => "title,gridProperties(rowCount,columnCount)"
                     ],
                  ]
               ]
             ]
             @session.request(:post, self.batch_update_url , :data => propupd.to_json, :header => Hash[ "Content-Type" => "application/json;charset=utf-8"], :response_type => :json )
             @meta_modified = false
             sent = true
           end

           if !@modified.empty?
             chunk = @modified
             #@modified.each_slice(400) do |chunk|

                data = Hash[
                  "valueInputOption" => "USER_ENTERED",
                  "data" => []
                ]

                for row, col in chunk
                   value = @cells[[row, col]]

                   data["data"] << Hash[
                      "range" => cell_name( row, col ),
                      "majorDimension" => "ROWS",
                      "values" => [[ value ]]
                   ]
                end
                @session.request(:post, self.batch_values_url , :data => data.to_json, :header => Hash[ "Content-Type" => "application/json;charset=utf-8"], :response_type => :json )
             #end



             @modified.clear()
             sent = true

           end

           return sent
        end

        # Calls save() and reload().
        def synchronize()
          save()
          reload()
        end

        # Deletes this worksheet. Deletion takes effect right away without calling save().
        def delete()
           raise "Doesn't work"
          ws_doc = @session.request(:get, self.worksheet_feed_url)
          edit_url = ws_doc.css("link[rel='edit']")[0]["href"]
          @session.request(:delete, edit_url)
        end

        # Returns true if you have changes made by []= which haven't been saved.
        def dirty?
          return !@modified.empty?
        end


        # Returns a [row, col] pair for a cell name string.
        # e.g.
        #   worksheet.cell_name_to_row_col("C2")  #=> [2, 3]
        def cell_name_to_row_col(cell_name)
          if !cell_name.is_a?(String)
            raise(ArgumentError, "Cell name must be a string: %p" % cell_name)
          end
          if !(cell_name.upcase =~ /^([A-Z]+)(\d+)$/)
            raise(ArgumentError,
                "Cell name must be only letters followed by digits with no spaces in between: %p" %
                    cell_name)
          end
          col = 0
          $1.each_byte() do |b|
            # 0x41: "A"
            col = col * 26 + (b - 0x41 + 1)
          end
          row = $2.to_i()
          return [row, col]
        end

        def inspect
          #fields = {:worksheet_feed_url => self.worksheet_feed_url}
          fields = {:worksheet_feed_url => self.worksheet_feed_url}
          fields[:title] = @title if @title
          return "\#<%p %s>" % [self.class, fields.map(){ |k, v| "%s=%p" % [k, v] }.join(", ")]
        end

      private

        def parse_cell_args(args)
          if args.size == 1 && args[0].is_a?(String)
            return cell_name_to_row_col(args[0])
          elsif args.size == 2 && args[0].is_a?(Integer) && args[1].is_a?(Integer)
            if args[0] >= 1 && args[1] >= 1
              return args
            else
              raise(ArgumentError,
                  "Row/col must be >= 1 (1-origin), but are %d/%d" % [args[0], args[1]])
            end
          else
            raise(ArgumentError,
                "Arguments must be either one String or two Integer's, but are %p" % [args])
          end
        end

    end

end
