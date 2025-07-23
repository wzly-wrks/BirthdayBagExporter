#!/bin/bash

echo "Birthday Bag Exporter - Installation and Launcher"
echo "================================================"
echo

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python 3 is not installed or not in PATH."
    echo "Please install Python 3.6 or higher from https://www.python.org/downloads/"
    echo "or use your system's package manager."
    echo
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Installing required packages..."
python3 -m pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "Failed to install required packages."
    echo "Please try running: pip3 install -r requirements.txt"
    echo
    read -p "Press Enter to exit..."
    exit 1
fi

echo
echo "Starting Birthday Bag Exporter..."
echo
python3 birthday_bag_exporter.py
if [ $? -ne 0 ]; then
    echo "Application exited with an error."
    echo
    read -p "Press Enter to exit..."
fi

exit 0