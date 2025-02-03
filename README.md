# QuantumShield

## Overview

QuantumShield is a Python-based tool designed to facilitate efficient data transfer and management across Windows applications. This program aims to improve workflow by automating file transfers between directories and enabling quick access to files within commonly used Windows applications like Microsoft Word.

## Features

- **Automated File Transfer**: Move files from a source directory to a destination directory seamlessly.
- **Integration with Microsoft Word**: Open transferred files directly in Microsoft Word for easy editing.
- **File Management**: List files in the destination directory to stay organized.

## Requirements

- Python 3.x
- `pywin32` package for Windows COM interface.
  
  Install it via pip:
  ```bash
  pip install pywin32
  ```

## Installation

1. Clone this repository to your local machine.
2. Ensure you have the required Python version and dependencies installed.
3. Adjust the `source` and `destination` directories in the `__main__` section of `quantum_shield.py` to match your system's file paths.

## Usage

1. Run the script via Python:
   ```bash
   python quantum_shield.py
   ```

2. The script will transfer all files from the specified source directory to the destination directory.
3. It will then list all files in the destination directory.
4. Modify the `open_in_word` method call with the correct filename to open a specific file in Microsoft Word.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For questions or suggestions, please open an issue in this repository.

```

Note: Make sure to adjust the file paths in the code to match your Windows environment, and ensure that Microsoft Word is installed on your system for the Word integration feature to work.