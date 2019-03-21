from os.path import dirname, join, abspath

# Base dir of project, contains subdirs "tests" and "oletools" and README.md
PROJECT_ROOT = dirname(dirname(dirname(abspath(__file__))))

# Directory with test data, independent of current working directory
DATA_BASE_DIR = join(PROJECT_ROOT, 'tests', 'test-data')

# Directory with source code
SOURCE_BASE_DIR = join(PROJECT_ROOT, 'oletools')
