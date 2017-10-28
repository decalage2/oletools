from .output_capture import OutputCapture

from os.path import dirname, join

# Directory with test data, independent of current working directory
DATA_BASE_DIR = join(dirname(dirname(__file__)), 'test-data')
