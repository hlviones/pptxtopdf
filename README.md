# PowerPoint to PDF Converter

This is a Python script that converts PowerPoint files (.ppt, .pptx) to PDF format. It uses the comtypes library to interact with Microsoft PowerPoint.

## Prerequisites

- Python 3.6 or above


## Dependancies 

- comtypes library (Install using `pip install comtypes`)

## Usage

Run the script with the following command:

- `input-folder`: Path to the folder containing the PowerPoint files to convert.
- `output-folder`: Path to the folder where the converted PDF files will be saved.

## Example

python pptxtopdf.py ./input ./output
This command will convert all the PowerPoint files in the `./input` folder and save the converted PDF files in the `./output` folder.

## Functionality

The script performs the following steps:

1. Validates the input folder path and checks if it exists.
2. Creates the output folder if it does not already exist.
3. Lists all the files in the input folder.
4. Iterates over each file and checks if it has a PowerPoint extension.
5. Creates a PowerPoint application object using comtypes.
6. Opens the PowerPoint slides.
7. Retrieves the base file name.
8. Constructs the output file path with the PDF extension.
9. Checks if the output file already exists and skips the conversion if it does.
10. Saves the slides as a PDF file.
11. Closes the slide deck.
12. Quits PowerPoint.
13. Keeps track of the successful conversions and errors.
14. Prints a summary of the conversion process at the end.

Note: The script runs without a user interface (`WithWindow=False`) to perform the conversions silently.

## Error Handling

If any error occurs during the conversion process, the script will catch the exception and print an error message specifying the file that caused the error.

## License

This script is released under the MIT License. Feel free to modify and use it according to your needs.
