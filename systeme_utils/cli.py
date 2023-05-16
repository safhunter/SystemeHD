import os
from fnmatch import fnmatch
from systeme_utils import JSON, SHOW, PLATFORM
from systeme_utils.systeme_utils import JsonConverter, PlatformConverter
class App:
    def __init__(self, name, version):
        self.name, self.version = name, version

    def __call__(self, args):
        if not args['command']:
            raise ValueError

        if args['command'] == JSON:
            converter = JsonConverter(os.path.dirname(args['filename']), os.path.basename(args['filename']))
            converter.convert()
        if args['command'] == PLATFORM:
            if args['batch']:
                root = args['batch']
                pattern = "*.xls"
                for path, subdirs, files in os.walk(root):
                    for name in files:
                        if fnmatch(name, pattern):
                            try:
                                converter = PlatformConverter(path, name)
                                converter.convert()
                            except Exception as ex:
                                print(f" Can't convert file {os.path.join(path, name)} cause:\n{ex}")
            else:
                converter = PlatformConverter(os.path.dirname(args['filename']), os.path.basename(args['filename']))
                converter.convert()
        elif args['command'] == SHOW:
             pass

