# Birthday Bag Exporter

A professional application for transforming Client Track exported Happy Birthday labels into a formatted route assignment spreadsheet.

![Birthday Bag Exporter](icon.png)

## Features

- **Modern, Professional Interface**: Clean design with intuitive controls
- **Dark Mode Support**: Toggle between light and dark themes
- **Drag & Drop Support**: Simply drag your Excel file onto the application
- **Automatic Package Installation**: Required packages are installed automatically
- **Route Assignment Editor**: Easily edit van numbers for routes, sorted by van number
- **Progress Tracking**: Visual progress bar shows processing status
- **Proper Formatting**: Creates Excel files with black separator bars between days

## Installation

1. Make sure you have Python 3.6+ installed
2. Download all files to a directory    
3. Run the appropriate installer for your system:

Windows: Double-click install_and_run_birthday_bag.bat

Mac/Linux: Run chmod +x install_and_run_birthday_bag.sh then ./install_and_run_birthday_bag.sh

4. Run the application:

```
python birthday_bag_exporter.py
```

The application will automatically install any required packages on first run.

## Usage

1. **Launch the application**:
   ```
   python birthday_bag_exporter.py
   ```

2. **Choose your theme**:
   - Toggle the "Dark Mode" checkbox in the top-right corner to switch between light and dark themes

3. **Process a file**:
   - Drag and drop your Happy Birthday labels Excel file onto the application
   - Or click "Browse" to select your file
   - Click "Process File" to generate the formatted output

4. **Edit Route Assignments** (if needed):
   - Click "Edit Route Assignments"
   - Routes are organized by day and sorted by van number
   - Update van numbers as needed
   - Add new routes using the form at the bottom of each tab
   - Click "Save Changes" when done

## Requirements

- Python 3.6+
- Required packages (automatically installed):
  - pandas
  - openpyxl
  - tkinterdnd2 (for drag & drop support)
  - pillow (for icon generation)

## File Structure

- `birthday_bag_exporter.py` - Main application
- `icon.png` / `icon.ico` - Application icons
- `requirements.txt` - Package requirements
- `README.md` - This documentation file
- `install_and_run_birthday_bag.bat` - Windows installer and launcher
- `install_and_run_birthday_bag.sh` - Mac/Linux installer and launcher

## Customization

The route assignments are stored in a dictionary within the application. You can edit them through the UI or directly in the code if needed.
