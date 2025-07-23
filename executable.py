# streamlit_app.py
import streamlit as st
import subprocess
import sys

subprocess.run([f"{sys.executable}", "cortaPgsPDF.py"])
