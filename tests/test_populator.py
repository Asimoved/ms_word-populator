# -*- coding: utf-8 -*-

import os
from ..populator import populate_word_form, list_placeholders, parse_filename_from_path, copy_file_with_timestamp, copy_and_populate_template

script_dir = os.path.dirname(__file__)
# test file
word_file_path = '../test_files/Hello_name.docx'
test_file_path = os.path.join(script_dir, word_file_path)
test_file_path = os.path.normpath(test_file_path)

# test sandbox folder
sandbox_rel_file_path = '../test_sandbox'
sandbox_file_path = os.path.join(script_dir, sandbox_rel_file_path)
sandbox_file_path = os.path.normpath(sandbox_file_path)

test_name = raw_input("What is your name?")
test_mood = raw_input("How are you feeling?")

dict_data = {
    'name': test_name,
    'mood': test_mood,
    }

copy_and_populate_template(test_file_path, dict_data, sandbox_file_path)
