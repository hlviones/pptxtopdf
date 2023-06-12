import sys
import os
import comtypes.client
import argparse


def convert(input_path, output_folder_path):
    if os.path.isdir(input_path):
        # Input path is a directory
        input_folder_path = os.path.abspath(input_path)

        if not os.path.isdir(input_folder_path):
            print("Error: Input folder does not exist.")
            return

        input_file_paths = [os.path.join(input_folder_path, file_name) for file_name in os.listdir(input_folder_path)]
    else:
        # Input path is a file
        input_file_paths = [os.path.abspath(input_path)]

    # Use the input_file_path's directory if output_folder_path is not provided
    if not output_folder_path:
        output_folder_path = os.path.dirname(input_file_paths[0])

    output_folder_path = os.path.abspath(output_folder_path)

    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    success_count = 0
    error_count = 0

    for input_file_path in input_file_paths:
        if not input_file_path.lower().endswith((".ppt", ".pptx")):
            print(f"Skipping file '{input_file_path}' as it does not have a PowerPoint extension.")
            continue

        try:
            # Create PowerPoint application object
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

            # Open the PowerPoint slides
            slides = powerpoint.Presentations.Open(input_file_path, WithWindow=False)

            # Get base file name
            file_name = os.path.splitext(os.path.basename(input_file_path))[0]

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
            print(f"Error converting file '{input_file_path}': {str(e)}")
            error_count += 1

    print(f"Conversion completed: {success_count} files converted successfully, {error_count} files failed.")


def main():
    parser = argparse.ArgumentParser(description='Convert PowerPoint files to PDF.')
    parser.add_argument('--input_dir', help='Path to the input folder containing PowerPoint files')
    parser.add_argument('--input_file', help='Path to the input PowerPoint file')
    parser.add_argument('--output_dir', help='[Optional] Path to the output folder to save the converted PDF files. '
                                              'If not provided, it defaults to the same directory as the input file '
                                              'or directory.')
    args = parser.parse_args()

    # Check if no arguments are provided
    if not any(vars(args).values()):
        parser.print_help()
        sys.exit(1)

    input_path = None
    if args.input_dir and args.input_file:
        print("Error: Please provide either --input_dir or --input_file, not both.")
        sys.exit(1)
    elif args.input_dir:
        input_path = args.input_dir
    elif args.input_file:
        input_path = args.input_file
    else:
        print("Error: Please provide either --input_dir or --input_file.")
        sys.exit(1)

    output_folder_path = args.output_dir  # Use the output_dir argument if provided


    convert(input_path, output_folder_path)


if __name__ == '__main__':
    main()
