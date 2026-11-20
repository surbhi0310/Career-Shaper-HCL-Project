import os
import sys

# Ensure Invoice_Data_Validation_Analysis is in sys.path
base_dir = os.path.join(os.path.dirname(__file__), 'Invoice_Data_Validation_Analysis')
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

# Import and execute the main application
from app import main
import streamlit as st

if __name__ == '__main__':
    main()
