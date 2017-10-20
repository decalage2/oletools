import os

def read(relative_path):
    test_data = os.path.dirname(os.path.abspath(__file__)) + '/test-data/'
    return open(test_data + relative_path, 'rb').read()

