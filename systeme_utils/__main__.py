"""SystemeHD Config converter entry point script."""
# systeme_utils/__main__.py

import argparse
import json

from systeme_utils import cli, __app_name__, __version__, COMMANDS

pretty_commands = json.dumps(COMMANDS, indent=4, sort_keys=True)
commands_list = list(COMMANDS.keys())
parser = argparse.ArgumentParser(prog=__app_name__,
                                 formatter_class=argparse.RawTextHelpFormatter,
                                 description=f'Provide some utils for SystemeHD PLC.\n'
                                             f'Supplied commands:\n'
                                             f'{commands_list}')
parser.add_argument('-c', '--command', choices=commands_list, required=True,
                    help=f'{pretty_commands}')
parser.add_argument('-f', '--filename',
                    help='Required file full name')
parser.add_argument('-b', '--batch',
                    help="Path to directory. Finds all *xls files in this dir and sub dirs, and tries to convert it")

args = parser.parse_args()


def main():
    app = cli.App(__app_name__, __version__)
    app(vars(args))


if __name__ == "__main__":
    main()
