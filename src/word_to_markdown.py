#!/usr/bin/env python3
"""
Convert Word documents to Markdown format.
"""

import logging
from pathlib import Path
from docx import Document  # This should import from python-docx package
import argparse

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def convert_word_to_markdown(input_path: str, output_path: str) -> None:
    """
    Convert a Word document to Markdown format.

    This function reads a Word document, processes its content, and converts it to Markdown format.
    It handles headings, bullet points, and text formatting (bold and italic).
    Args:
        input_path (str): The file path of the input Word document to be converted.
        output_path (str): The file path where the converted Markdown document will be saved.

    Returns:
        None

    Raises:
        Exception: If any error occurs during the conversion process.

    Note:
        The function logs the success or failure of the conversion process.
    """
    try:
        # Load the Word document
        doc = Document(input_path)
        markdown_content = []

        # Process each paragraph
        for paragraph in doc.paragraphs:
            # Skip empty paragraphs
            if not paragraph.text.strip():
                continue

            text = paragraph.text
            style = paragraph.style.name

            # Convert heading styles
            if style.startswith('Heading'):
                level = style[-1]  # Get the heading level (1-9)
                markdown_content.append(f"{'#' * int(level)} {text}\n")

            # Handle bullet points
            elif style.startswith('List'):
                markdown_content.append(f"* {text}\n")

            # Regular paragraphs
            else:
                # Check for bold and italic text
                runs = paragraph.runs
                paragraph_content = []

                for run in runs:
                    text = run.text
                    if run.bold and run.italic:
                        text = f"***{text}***"
                    elif run.bold:
                        text = f"**{text}**"
                    elif run.italic:
                        text = f"*{text}*"
                    paragraph_content.append(text)

                markdown_content.append(''.join(paragraph_content) + '\n')

        # Write the markdown content to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(markdown_content))

        logger.info(f"Successfully converted {input_path} to {output_path}")

    except Exception as e:
        logger.error(f"Error converting document: {str(e)}", exc_info=True)
        raise


def main():
    """
    Main function to run the Word to Markdown converter.

    This function sets up the argument parser, processes command-line arguments,
    and calls the conversion function. It handles the input file path and
    determines the output file path.

    Parameters:
    None

    Returns:
    int: Returns 0 if the conversion is successful, 1 if an error occurs.

    Raises:
    Exception: Any exception that occurs during the execution is caught and logged.
    """
    parser = argparse.ArgumentParser(
        description='Convert Word documents to Markdown format.'
    )
    parser.add_argument(
        'input_file',
        help='Path to the input Word document'
    )
    parser.add_argument(
        '-o', '--output',
        help='Path for the output Markdown file (default: input_file_name.md)',
        default=None
    )

    try:
        args = parser.parse_args()

        # If no output file is specified, use input filename with .md extension
        output_file = args.output
        if output_file is None:
            output_file = Path(args.input_file).with_suffix('.md')

        convert_word_to_markdown(args.input_file, output_file)
        return 0

    except Exception as e:
        logger.error(f"An error occurred: {str(e)}")
        return 1


if __name__ == "__main__":
    exit(main())