# Microsoft Word Edit Time Tracker

## Overview

Word Edit Time Tracker is a Python script designed to calculate the total editing time for all Word documents in a specified folder. It helps users monitor their productivity and manage their time effectively.

## Features

- Calculate total editing time for Word documents in a folder
- Display editing time for each document
- Summarize total editing time in hours and minutes

## Requirements

- Python 3.x
- `pywin32` library

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/GlennAntonySheen/Microsoft-Word-Edit-Time-Tracker.git
   ```
2. Navigate to the project directory:
   ```bash
   cd Microsoft-Word-Edit-Time-Tracker
   ```
3. Install the required dependencies:
   ```bash
   pip install pywin32
   ```

## Usage

1. Set the folder path where your Word files are located by modifying the `folder_path` variable in `main.py`:
   ```python
   folder_path = r"your\path\to\word\files"
   ```
2. Run the script:
   ```bash
   python main.py
   ```

## Example Output

```
document1.docx: 15 minutes
document2.docx: 30 minutes
...

Total Editing Time for all documents: 1 hr 45 minutes
```

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request.

## License

This project is licensed under the MIT License.
