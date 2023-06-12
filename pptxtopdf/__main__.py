import sys
import os
import comtypes.client

def convert(input_folder_path, output_folder_path):
    input_folder_path = os.path.abspath(input_folder_path)
    output_folder_path = os.path.abspath(output_folder_path)

    if not os.path.isdir(input_folder_path):
        print("Error: Input folder does not exist.")
        return

    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    input_file_paths = os.listdir(input_folder_path)
    success_count = 0
    error_count = 0

    for input_file_name in input_file_paths:
        # Skip if file does not contain a PowerPoint extension
        if not input_file_name.lower().endswith((".ppt", ".pptx")):
            continue

        # Create input file path
        input_file_path = os.path.join(input_folder_path, input_file_name)

        try:
            # Create PowerPoint application object
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

            # Open the PowerPoint slides
            slides = powerpoint.Presentations.Open(input_file_path, WithWindow=False)

            # Get base file name
            file_name = os.path.splitext(input_file_name)[0]

            # Create output file path
            output_file_path = os.path.join(output_folder_path, file_name + ".pdf")

            if os.path.exists(output_file_path):
                print(f"Error: Output file '{output_file_path}' already exists.")
                error_count += 1
                continue

            # Save as PDF (formatType = 32)
            slides.SaveAs(output_file_path, 32)

            # Close the slide deck
            slides.Close()

            powerpoint.Quit()

            success_count += 1
        except Exception as e:
            print(f"Error converting file '{input_file_name}': {str(e)}")
            error_count += 1

    print(f"Conversion completed: {success_count} files converted successfully, {error_count} files failed.")


def main():
    if len(sys.argv) != 3:
        print("Usage: pptxtopdf input-folder output-folder")
    else:
        input_folder_path = sys.argv[1]
        output_folder_path = sys.argv[2]
        convert(input_folder_path, output_folder_path)

if __name__ == '__main__':
    main()