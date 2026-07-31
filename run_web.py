#!/usr/bin/env python
"""
Launcher for the Microsoft Word Page Exporter Pro Web Interface
"""
import sys
import os

# Add root directory to sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from word_exporter_pro.web.app import app, main

application = app

if __name__ == "__main__":
    main()
