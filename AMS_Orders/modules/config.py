"""
Configuration loader for AOMOSO Download Manager.

Reads settings from config.json in the working directory.
If config.json is not found, falls back to config.example.json
from the project root with a warning.
"""

import json
import os
import sys

_config = None


def _find_project_root():
    """Walk up from the script/exe location to find the project root
    (directory containing config.example.json or config.json)."""
    if getattr(sys, 'frozen', False):
        search_dir = os.path.dirname(sys.executable)
    else:
        search_dir = os.getcwd()

    # Check current dir first, then walk up at most 3 levels
    for _ in range(4):
        if os.path.exists(os.path.join(search_dir, 'config.json')):
            return search_dir
        if os.path.exists(os.path.join(search_dir, 'config.example.json')):
            return search_dir
        search_dir = os.path.dirname(search_dir)

    return os.getcwd()


def _load_config():
    """Load configuration from config.json, falling back to config.example.json."""
    global _config
    if _config is not None:
        return _config

    root = _find_project_root()
    config_path = os.path.join(root, 'config.json')
    example_path = os.path.join(root, 'config.example.json')

    if os.path.exists(config_path):
        with open(config_path, 'r') as f:
            _config = json.load(f)
    elif os.path.exists(example_path):
        import warnings
        warnings.warn(
            "config.json not found — using config.example.json defaults. "
            "Copy config.example.json to config.json and fill in your values.",
            stacklevel=2,
        )
        with open(example_path, 'r') as f:
            _config = json.load(f)
    else:
        raise FileNotFoundError(
            "Neither config.json nor config.example.json found. "
            "Please create config.json from the config.example.json template."
        )

    return _config


def get_sap_config():
    """Return the SAP configuration dict."""
    return _load_config()["sap"]


def get_web_config():
    """Return the web configuration dict."""
    return _load_config()["web"]
