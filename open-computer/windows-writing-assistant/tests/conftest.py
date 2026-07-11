import os
import sys

# Make the docd package importable when pytest runs from any directory.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
