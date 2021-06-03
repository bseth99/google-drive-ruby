# Author: Guy Boertje <https://github.com/guyboertje>
# Author: David R. Albrecht <https://github.com/eldavido>
# Author: Hiroshi Ichikawa <http://gimite.net/>
# The license of this source is "New BSD Licence"

require "google_drive_v3/acl_entry"

module GoogleDriveV3

    # ACL (access control list) of a spreadsheet.
    #
    # Use GoogleDrive::Spreadsheet#acl to get GoogleDrive::Acl object.
    # See GoogleDrive::Spreadsheet#acl for usage example.
    #
    # This code is based on https://github.com/guyboertje/gdata-spreadsheet-ruby .
    class Acl

        include(Util)
        extend(Forwardable)

        def initialize(session, file) #:nodoc:
          @session = session
          @file = file
          api_result = @session.execute!(
              :api_method => @session.drive.permissions.list,
              :parameters => { "fileId" => @file.id })
          @entries = api_result.data.items.map(){ |i| AclEntry.new(i, self) }
        end

        def_delegators(:@entries, :size, :[], :each)

        # Adds a new entry. +entry+ is either a GoogleDrive::AclEntry or a Hash with keys
        # :scope_type, :scope and :role. See GoogleDrive::AclEntry#scope_type and
        # GoogleDrive::AclEntry#role for the document of the fields.
        #
        # NOTE: This sends email to the new people.
        #
        # e.g.
        #   # A specific user can read or write.
        #   spreadsheet.acl.push(
        #       {:scope_type => "user", :scope => "example2@gmail.com", :role => "reader"})
        #   spreadsheet.acl.push(
        #       {:scope_type => "user", :scope => "example3@gmail.com", :role => "writer"})
        #   # Publish on the Web.
        #   spreadsheet.acl.push(
        #       {:scope_type => "default", :role => "reader"})
        #   # Anyone who knows the link can read.
        #   spreadsheet.acl.push(
        #       {:scope_type => "default", :with_key => true, :role => "reader"})
        def push(params)
          new_permission = @session.drive.permissions.insert.request_schema.new(
              convert_params(params))
          api_result = @session.execute!(
              :api_method => @session.drive.permissions.insert,
              :body_object => new_permission,
              :parameters => { "fileId" => @file.id })
          new_entry = AclEntry.new(api_result.data, self)
          @entries.push(new_entry)
          return new_entry
        end

        # Deletes an ACL entry.
        #
        # e.g.
        #   spreadsheet.acl.delete(spreadsheet.acl[1])
        def delete(entry)
          @session.execute!(
              :api_method => @session.drive.permissions.delete,
              :parameters => {
                  "fileId" => @file.id,
                  "permissionId" => entry.id,
                })
          @entries.delete(entry)
        end

        def update_role(entry) #:nodoc:
          api_result = @session.execute!(
              :api_method => @session.drive.permissions.update,
              :body_object => entry.api_permission,
              :parameters => {
                  "fileId" => @file.id,
                  "permissionId" => entry.id,
              })
          entry.api_permission = api_result.data
          return entry
        end

        def inspect
          return "\#<%p %p>" % [self.class, @entries]
        end

      private

        def convert_params(orig_params)
          new_params = {}
          for k, v in orig_params
            k = k.to_s()
            case k
              when "scope_type"
                new_params["type"] = (v == "default" ? "anyone" : v)
              when "scope"
                new_params["value"] = v
              when "with_key"
                new_params["withLink"] = v
              else
                new_params[k] = v
            end
          end
          return new_params
        end

    end

end
