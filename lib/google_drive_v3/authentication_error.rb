# Author: Hiroshi Ichikawa <http://gimite.net/>
# The license of this source is "New BSD Licence"

require "google_drive_v3/error"


module GoogleDriveV3

    # Raised when GoogleDrive.login has failed.
    class AuthenticationError < GoogleDriveV3::Error

    end

end
