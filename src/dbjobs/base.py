import sys

class Base(object):
    """Base class for project"""

    def error(self, msg):
        sys.exit(f"ERROR: {msg}")

    def warning(self, msg):
        print(f"WARNING: {msg}")

    def debug(self, msg):
        if self.debug_opt:
            print(f"DEBUG: {msg}")
