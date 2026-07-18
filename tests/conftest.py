"""Hace importable el codigo del proyecto desde los tests."""
import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
