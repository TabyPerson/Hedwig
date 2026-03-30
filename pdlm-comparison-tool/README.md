# pdlm-comparison-tool

## Overview
The PDLM Comparison Tool is designed to facilitate the comparison of various documentation types, including requirements, protocols, records, and defects. This tool provides a user-friendly interface for selecting different comparison options and dynamically executes the corresponding scripts.

## Project Structure
```
pdlm-comparison-tool
├── src
│   ├── main.py                     # Entry point for the application
│   ├── options                     # Contains scripts for different comparison options
│   │   ├── urs_tm_comparison.py    # Comparison between URS documents and TM APP files
│   │   ├── tm_val_tp_comparison.py  # Comparison between TM APP files and Validation Test Protocols
│   │   ├── val_tp_records_comparison.py # Comparison between Validation Test Protocols and Records
│   │   ├── val_tm_records_comparison.py # Comparison between Validation TM APP files and Records
│   │   ├── val_records_pdsr_comparison.py # Comparison between Validation Test Records and PDSR
│   │   ├── product_validation_report.py # Generates a comprehensive Product Validation Report
│   │   └── check_video.py          # Checks for missing or invalid video/file paths
│   ├── utils                        # Utility functions for file and comparison operations
│   │   ├── file_utils.py           # File handling operations
│   │   ├── comparison_utils.py      # Helper functions for comparisons
│   │   └── __init__.py             # Marks the utils directory as a package
│   └── __init__.py                 # Marks the src directory as a package
├── requirements.txt                 # Lists project dependencies
└── README.md                        # Documentation for the project
```

## Installation
To set up the project, clone the repository and install the required dependencies:

```bash
git clone <repository-url>
cd pdlm-comparison-tool
pip install -r requirements.txt
```

## Usage
Run the application by executing the main script:

```bash
python src/main.py
```

Follow the prompts to select the desired comparison option. The tool will guide you through the process of uploading files and executing the comparison.

## Comparison Options
The following comparison options are available:

1. **URS DOC x TM APP Comparison**: Compare requirements from URS documents with TM APP files.
2. **TM APP x Validation Test Protocol Comparison**: Compare TM APP files with Validation Test Protocols.
3. **Validation Test Protocol x Records Comparison**: Compare Validation Test Protocols with Validation Test Records.
4. **Validation TM APP x Validation Test Records Comparison**: Compare Validation TM APP files with Validation Test Records.
5. **Validation Test Records x PDSR Comparison**: Compare Validation Test Records with Product Defect Status Reports.
6. **Product Validation Report**: Generate a comprehensive report by comparing multiple files.
7. **Validation - Check Video**: Analyze video evidence paths for missing or invalid entries.

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.