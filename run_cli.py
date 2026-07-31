#!/usr/bin/env python
"""
Launcher for Microsoft Word Page Exporter Pro CLI
"""
import sys
import os

# Add root directory to sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from word_exporter_pro.cli import main

if __name__ == "__main__":
    main()
