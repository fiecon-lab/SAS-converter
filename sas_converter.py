import pandas as pd
import pyreadstat
from pathlib import Path
import logging
import argparse

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def generate_output_path(input_path, output_path=None):
    """
    Generate output path with format suffix in filename
    """
    input_path = Path(input_path)
    original_format = input_path.suffix.lstrip('.')
    stem = input_path.stem
    
    if output_path is None:
        # Generate name like "filename-xpt.xlsx" or "filename-sas7bdat.xlsx"
        new_filename = f"{stem}-{original_format}.xlsx"
        return str(input_path.parent / new_filename)
    else:
        # If output_path is a directory, create the new filename inside it
        output_path = Path(output_path)
        if output_path.is_dir() or not output_path.suffix:
            new_filename = f"{stem}-{original_format}.xlsx"
            return str(output_path / new_filename)
        return str(output_path)

def convert_sas7bdat_to_excel(input_path, output_path=None):
    """
    Convert .sas7bdat file to Excel format (.xlsx)
    """
    try:
        # Read the SAS file
        df, meta = pyreadstat.read_sas7bdat(input_path)
        
        # Generate output path with format suffix
        output_path = generate_output_path(input_path, output_path)
        
        # Create output directory if it doesn't exist
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        # Save to Excel
        df.to_excel(output_path, index=False)
        logger.info(f"Successfully converted {input_path} to {output_path}")
        
        return output_path
        
    except Exception as e:
        logger.error(f"Error converting {input_path}: {str(e)}")
        raise

def convert_xpt_to_excel(input_path, output_path=None):
    """
    Convert .xpt file to Excel format (.xlsx)
    """
    try:
        # Read the XPT file using pandas' built-in reader
        df = pd.read_sas(input_path, format='xport')
        
        # Generate output path with format suffix
        output_path = generate_output_path(input_path, output_path)
        
        # Create output directory if it doesn't exist
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        
        # Save to Excel
        df.to_excel(output_path, index=False)
        logger.info(f"Successfully converted {input_path} to {output_path}")
        
        return output_path
        
    except Exception as e:
        logger.error(f"Error converting {input_path}: {str(e)}")
        raise

def convert_folder(input_directory, output_directory=None):
    """
    Convert all SAS files in a directory to Excel format
    """
    input_path = Path(input_directory)
    if not input_path.exists():
        raise FileNotFoundError(f"Input directory '{input_directory}' does not exist")
    
    output_path = Path(output_directory) if output_directory else input_path
    
    # Create output directory if it doesn't exist
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Keep track of processed files
    processed_files = 0
    successful_conversions = 0
    
    # Process all .sas7bdat files
    for file in input_path.glob('*.sas7bdat'):
        processed_files += 1
        try:
            convert_sas7bdat_to_excel(str(file), str(output_path))
            successful_conversions += 1
        except Exception as e:
            logger.error(f"Failed to convert {file}: {str(e)}")
    
    # Process all .xpt files
    for file in input_path.glob('*.xpt'):
        processed_files += 1
        try:
            convert_xpt_to_excel(str(file), str(output_path))
            successful_conversions += 1
        except Exception as e:
            logger.error(f"Failed to convert {file}: {str(e)}")
    
    # Log summary
    logger.info(f"\nConversion Summary:")
    logger.info(f"Total files processed: {processed_files}")
    logger.info(f"Successful conversions: {successful_conversions}")
    logger.info(f"Failed conversions: {processed_files - successful_conversions}")
    
    if processed_files == 0:
        logger.warning("No .sas7bdat or .xpt files found in the specified directory")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Convert all SAS files in a directory to Excel format')
    parser.add_argument('input_folder', help='Input folder containing SAS files')
    parser.add_argument('--output', help='Output directory (optional, defaults to input folder)')
    
    args = parser.parse_args()
    
    try:
        convert_folder(args.input_folder, args.output)
    except Exception as e:
        logger.error(f"Conversion failed: {str(e)}")
        exit(1)