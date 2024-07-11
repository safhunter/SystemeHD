"""Top-level package for SystemeHD Config converter."""
# systeme_utils/__init__.py

__app_name__ = "systeme_utils"
__version__ = "0.2.0"

JSON = 'json'
NEW_JSON = 'new-json'
PLATFORM = 'platform'
SHOW = 'show'

COMMANDS = {
    JSON: "Converts SystemHD old config *.xls file to *.json. Uses the file provided by --filename option",
    NEW_JSON: "Converts SystemHD new config *.xls file to *.json. Uses the file provided by --filename option",
    PLATFORM: "Converts SystemHD config *.xls file to Platform HD config *.xlsx. Uses the file provided by --filename "
              "option",
    SHOW: "Shows config *.json. Uses the file provided by --filename option"
}
