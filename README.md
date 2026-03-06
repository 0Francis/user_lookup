# User Lookup Project

## Overview
The User Lookup project is designed to process user data from an Excel file, extract relevant information, and facilitate the retrieval of user details from a domain. The project utilizes Python and various libraries to streamline data handling and command execution.

## Project Structure
```
UserLookup
├── src
│   ├── user_lookup_core.ipynb
│   └── utils
│       └── extract_staff_name.py
├── requirements.txt
└── README.md
```

## File Descriptions

- **src/user_lookup_core.ipynb**: This Jupyter Notebook contains the main logic for processing user data. It imports necessary libraries, reads an Excel file, processes the data to extract AUUID digits, and prepares for writing the results to a new sheet.

- **src/utils/extract_staff_name.py**: This Python script defines a function that executes the command `net user /domain "user_id"` for each AUUID digit to extract the full name of the user. It handles command execution and parsing of the output.

- **requirements.txt**: This file lists the dependencies required for the project, including `pandas` and `openpyxl`, which are essential for handling Excel files.

## Setup Instructions
1. Clone the repository to your local machine.
2. Navigate to the project directory.
3. Install the required dependencies using the following command:
   ```
   pip install -r requirements.txt
   ```

## Usage Guidelines
1. Open the `user_lookup_core.ipynb` file in Jupyter Notebook.
2. Modify the `RAW_FILE` variable to point to your Excel file containing user data.
3. Run the cells in the notebook to process the data and extract AUUID digits.
4. The results will be saved in a new sheet called "user data" in the specified Excel file.

## Contributing
Contributions to the project are welcome. Please submit a pull request or open an issue for any enhancements or bug fixes.