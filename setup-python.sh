#!/bin/bash
# Setup script for ClipSon Python version

echo "Installing system dependencies..."
sudo apt update
sudo apt install -y python3 python3-pip xclip libnotify-bin

echo "Installing Python dependencies..."
pip3 install -r requirements.txt

echo "Setup complete! Run with: python3 clipson.py"
