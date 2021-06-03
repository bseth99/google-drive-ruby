# Author: Hiroshi Ichikawa <http://gimite.net/>
# The license of this source is "New BSD Licence"

require "google_drive_v4/error"


module GoogleDriveV4

    # Raised when GoogleDrive.login has failed.
    class AuthenticationError < GoogleDriveV4::Error

    end

end
