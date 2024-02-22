#!/usr/bin/env python3

import os

script_dir = os.path.dirname(os.path.realpath(__file__))

os.chdir(script_dir)

print(os.getcwd())