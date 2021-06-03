# Author: Hiroshi Ichikawa <http://gimite.net/>
# The license of this source is "New BSD Licence"

require "time"

require "google_drive_v4/util"
require "google_drive_v4/error"
require "google_drive_v4/worksheet"
require "google_drive_v4/table"
require "google_drive_v4/acl"
require "google_drive_v4/file"


module GoogleDriveV4

    # A spreadsheet.
    #
    # Use methods in GoogleDrive::Session to get GoogleDrive::Spreadsheet object.
    class Spreadsheet < GoogleDriveV4::File

        include(Util)

        SUPPORTED_EXPORT_FORMAT = Set.new(["xls", "csv", "pdf", "ods", "tsv", "html"])

        def key
          return self.id
        end

        # URL of worksheet-based feed of the spreadsheet.
        def worksheets_feed_url params = Hash[]
           uri = URI("https://sheets.googleapis.com/v4/spreadsheets/%s" % self.id )
           uri.query = URI.encode_www_form( Hash[ "includeGridData" => false ].merge( params ) )
           return uri.to_s
        end

        def worksheet_url title
           return "https://sheets.googleapis.com/v4/spreadsheets/%s/values/%s" % [ self.id, URI.escape( title ) ]
        end

        def worksheets_data params=Hash[]
           @session.request(:get, self.worksheets_feed_url( params ), :header => Hash[ "Content-Type" => "application/json;charset=utf-8"], :response_type => :json )
        end

        # Returns worksheets of the spreadsheet as array of GoogleDrive::Worksheet.
        def worksheets
          doc = self.worksheets_data
          if !doc.is_a?( Hash ) && !doc.has_key?("sheets")
            raise(GoogleDriveV4::Error,
                "%s did not return a list of sheets." %
                self.worksheets_feed_url)
          end
          result = []
          doc["sheets"].each do |entry|
            title = entry["properties"]["title"]
            updated = self.api_file.modifiedDate
            url = self.worksheet_url title
            result.push(ws = Worksheet.new(@session, self, url, title, updated))
            ws.properties = entry["properties"]
          end
          return result.freeze()
        end

        # Returns a GoogleDrive::Worksheet with the given title in the spreadsheet.
        #
        # Returns nil if not found. Returns the first one when multiple worksheets with the
        # title are found.
        def worksheet_by_title(title)
          return self.worksheets.find(){ |ws| ws.title == title }
        end

        # Adds a new worksheet to the spreadsheet. Returns added GoogleDrive::Worksheet.
        def add_worksheet(title, max_rows = 100, max_cols = 20)
           raise "Doesn't work"
          xml = <<-"EOS"
            <entry xmlns='http://www.w3.org/2005/Atom'
                   xmlns:gs='http://schemas.google.com/spreadsheets/2006'>
              <title>#{h(title)}</title>
              <gs:rowCount>#{h(max_rows)}</gs:rowCount>
              <gs:colCount>#{h(max_cols)}</gs:colCount>
            </entry>
          EOS
          doc = @session.request(:post, self.worksheets_feed_url, :data => xml)
          url = doc.css(
            "link[rel='http://schemas.google.com/spreadsheets/2006#cellsfeed']")[0]["href"]
          return Worksheet.new(@session, self, url, title)
        end

    end

end
