import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import os
import argparse
import urllib.request
import ssl
import zipfile
import shutil
from pathlib import Path
from datetime import datetime

def ensure_data_folder(data_folder='data', force_refresh=False):
    """
    Ensure data folder exists with Excel files. If not, download and extract from gouv.fr.
    
    Parameters:
    data_folder (str): Path to the data folder (default: 'data')
    force_refresh (bool): Force download even if data exists (default: False)
    
    Returns:
    bool: True if data folder exists with files, False otherwise
    """
    data_path = Path(data_folder)
    
    # Check if data folder exists and has Excel files
    if data_path.exists() and not force_refresh:
        excel_files = list(data_path.glob('*.xlsx')) + list(data_path.glob('*.xls'))
        if excel_files:
            print(f"Found {len(excel_files)} Excel file(s) in '{data_folder}' folder.")
            return True
    
    # Need to download data
    print("Downloading pollen data from data.gouv.fr...")
    print("This may take a minute...")
    
    # Create data folder if it doesn't exist
    data_path.mkdir(parents=True, exist_ok=True)
    
    # The actual download URL (we need to get the direct link)
    # Based on gouv.fr API, we construct the direct download URL
    zip_url = "https://www.data.gouv.fr/api/1/datasets/r/d8c275e4-9e8b-4c58-97fe-8f0d48d2d5c7"
    
    zip_file = data_path / 'pollen_data.zip'
    
    try:
        # Create SSL context that bypasses certificate verification
        # (needed for some systems with certificate issues)
        ssl_context = ssl.create_default_context()
        ssl_context.check_hostname = False
        ssl_context.verify_mode = ssl.CERT_NONE
        
        # Download the zip file
        print(f"Downloading from: {zip_url}")
        with urllib.request.urlopen(zip_url, context=ssl_context) as response:
            with open(zip_file, 'wb') as out_file:
                out_file.write(response.read())
        print(f"Downloaded to: {zip_file}")
        
        # Extract the zip file
        print(f"Extracting files from zip...")
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            # Extract all files, flattening the folder structure
            for member in zip_ref.namelist():
                # Get just the filename without the directory path
                filename = os.path.basename(member)
                if filename:  # Skip if it's a directory
                    # Extract to data_path directly
                    source = zip_ref.open(member)
                    target_path = data_path / filename
                    with open(target_path, 'wb') as target:
                        target.write(source.read())
                    source.close()
        
        # Remove the zip file
        zip_file.unlink()
        print(f"Data extracted successfully to '{data_folder}' folder.")
        
        # Verify we have Excel files
        excel_files = list(data_path.glob('*.xlsx')) + list(data_path.glob('*.xls'))
        if excel_files:
            print(f"Found {len(excel_files)} Excel file(s) in '{data_folder}' folder.")
            return True
        else:
            print("Warning: No Excel files found after extraction.")
            return False
    
    except Exception as e:
        print(f"Error downloading or extracting data: {e}")
        print(f"Please manually download from: https://www.data.gouv.fr/datasets/donnees-historiques-de-surveillance-des-pollens-et-des-moisissures?resource_id=d8c275e4-9e8b-4c58-97fe-8f0d48d2d5c7")
        print(f"And extract the Excel files to the '{data_folder}' folder.")
        return False


def extract_alnus_data(folder_path, city_name='NICE', allergen_col=None):
    """
    Extract allergen data and date from all Excel files in a folder.
    
    Parameters:
    folder_path (str): Path to the folder containing Excel files
    city_name (str): City name to search for in filenames (default: 'NICE')
    allergen_col (int or str): Column index (0-based) or name for the allergen. If None, defaults to column 6 (G)
    
    Returns:
    pd.DataFrame: DataFrame with columns 'date', 'year', and 'allergen'
    """
    # keep first 8 charactes of the city_name to avoid issues with long city names
    city_name = city_name[:8]
    if allergen_col is None:
        allergen_col = 6  # Default to column G (index 6)
    
    data = []
    excel_files = list(Path(folder_path).glob(f'*{city_name}*.xlsx')) + list(Path(folder_path).glob(f'*{city_name}*.xls'))
    
    if not excel_files:
        print(f"No Excel files found in {folder_path} for city '{city_name}'")
        return pd.DataFrame()
    
    print(f"Found {len(excel_files)} Excel file(s)")
    
    for excel_file in excel_files:
        print(f"Processing: {excel_file.name}")
        try:
            # Read Excel file
            df = pd.read_excel(excel_file)
            
            # Convert allergen_col to integer if it's a column name
            if isinstance(allergen_col, str):
                if allergen_col in df.columns:
                    col_index = df.columns.get_loc(allergen_col)
                else:
                    print(f"Warning: Column '{allergen_col}' not found in {excel_file.name}. Skipping.")
                    continue
            else:
                col_index = allergen_col
            
            # Check if columns exist
            if len(df.columns) > max(0, col_index):
                df_subset = df.iloc[:, [0, col_index]].copy()  # Select columns A and allergen
                allergen_name = df.columns[col_index] if isinstance(df.columns[col_index], str) else f'Column {chr(65 + col_index)}'
                df_subset.columns = ['date', 'allergen']
                
                # Convert date column to datetime if needed
                df_subset['date'] = pd.to_datetime(df_subset['date'], errors='coerce')
                
                # Extract year from date
                df_subset['year'] = df_subset['date'].dt.year
                
                # Remove rows with missing values
                df_subset = df_subset.dropna()
                
                data.append(df_subset)
            else:
                print(f"Warning: Column index {col_index} does not exist in {excel_file.name}. Skipping.")
        
        except Exception as e:
            print(f"Error reading {excel_file.name}: {e}")
    
    if data:
        combined_df = pd.concat(data, ignore_index=True)
        return combined_df
    else:
        return pd.DataFrame()


def get_available_columns(folder_path, city_name='NICE'):
    """
    Detect available columns in Excel files.
    
    Parameters:
    folder_path (str): Path to the folder containing Excel files
    city_name (str): City name to search for in filenames
    
    Returns:
    list: List of available column names
    """
    # keep first 8 charactes of the city_name to avoid issues with long city names
    city_name = city_name[:8]
    excel_files = list(Path(folder_path).glob(f'*{city_name}*.xlsx')) + list(Path(folder_path).glob(f'*{city_name}*.xls'))
    
    if excel_files:
        try:
            df = pd.read_excel(excel_files[0])
            return list(df.columns[1:])  # Return all columns except the first (date)
        except Exception as e:
            print(f"Error reading {excel_files[0]}: {e}")
    
    return []


def plot_allergen_by_week(df, allergen_name='ALNUS', num_years=10, city_name='NICE', output_file='allergen_plot.png'):
    """
    Create a scatter plot with week on x-axis and allergen values on y-axis.
    
    Parameters:
    df (pd.DataFrame): DataFrame with columns 'date', 'year', and 'allergen'
    allergen_name (str): Name of the allergen to display in the title
    num_years (int): Number of years to plot (default: 10)
    city_name (str): Name of the city (default: 'NICE')
    output_file (str): Path to save the plot image
    """
    
    if df.empty:
        print("No data to plot")
        return
    
    # Filter to last N years
    max_year = df['year'].max()
    min_year = max_year - (num_years - 1)
    df_filtered = df[df['year'] >= min_year].copy()
    
    plt.figure(figsize=(14, 7))
    
    # Sort by date for better visualization
    df_sorted = df_filtered.sort_values('date').copy()
    
    # Group by year and week
    df_sorted['year_week'] = df_sorted['date'].dt.to_period('W')
    
    # Calculate mean allergen value per week
    weekly_mean = df_sorted.groupby('year_week')['allergen'].mean()
    
    # Convert period index to timestamp for better plotting
    week_labels = weekly_mean.index.to_timestamp()
    
    # Plot
    plt.scatter(week_labels, weekly_mean.values, s=100, alpha=0.6, color='blue', label=f'Mean {allergen_name}')
    
    # Optional: Add a line plot for better visualization
    plt.plot(week_labels, weekly_mean.values, alpha=0.3, color='blue')
    
    # Labels and title
    plt.xlabel('Week', fontsize=12, fontweight='bold')
    plt.ylabel(f'{allergen_name} Value', fontsize=12, fontweight='bold')
    plt.title(f'{allergen_name} Values in {city_name} ({min_year}-{max_year})', fontsize=14, fontweight='bold')
    plt.grid(True, alpha=0.3)
    plt.legend()
    
    # Format x-axis to show month - year format
    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%y'))
    
    # Rotate x-axis labels for better readability
    plt.xticks(rotation=45, ha='right')
    
    # Save the plot
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    print(f"Plot saved to: {output_file}")
    
    # Display the plot
    plt.show()


if __name__ == "__main__":
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description='Extract allergen data from Excel files and create a visualization.'
    )
    parser.add_argument(
        'folder_path',
        nargs='?',
        default=None,
        help='Path to the folder containing Excel files'
    )
    parser.add_argument(
        '-c', '--city',
        default='NICE',
        help='City name to search for in filenames (default: NICE)'
    )
    parser.add_argument(
        '-a', '--allergen',
        default=None,
        help='Allergen column name or index (0-based) to plot. If not found, a list of available allergens will be shown.'
    )
    parser.add_argument(
        '-y', '--years',
        type=int,
        default=10,
        help='Number of years to plot (default: 10)'
    )
    parser.add_argument(
        '-r', '--refresh',
        action='store_true',
        help='Force download and refresh data from data.gouv.fr'
    )
    
    args = parser.parse_args()
    
    # Ensure data folder has Excel files
    if not ensure_data_folder('data', args.refresh):
        print("Cannot proceed without data files.")
        exit(1)
    
    # Get folder path from argument or use default data folder
    folder_path = args.folder_path
    if not folder_path:
        folder_path = 'data'
        if not os.path.isdir(folder_path):
            folder_path = input("Enter the folder path containing Excel files: ").strip()
    
    if os.path.isdir(folder_path):
        # Check if allergen is specified as an index
        allergen_col = 6  # Default to column G (index 6)
        allergen_name = 'ALNUS'
        
        if args.allergen:
            try:
                # Try to parse as integer (column index)
                allergen_col = int(args.allergen)
                allergen_name = f'Column {chr(65 + allergen_col)}'
            except ValueError:
                # Use as column name
                allergen_col = args.allergen
                allergen_name = args.allergen
        
        # Extract data
        data_df = extract_alnus_data(folder_path, args.city, allergen_col)
        
        if data_df.empty:
            # Get available columns and suggest them
            print("\nTrying to detect available allergens...")
            available_cols = get_available_columns(folder_path, args.city)
            if available_cols:
                print("Available allergen columns:")
                for idx, col in enumerate(available_cols, start=1):
                    print(f"  {idx}. {col} (index: {idx})")
                print("\nPlease specify an allergen using -a/--allergen option.")
                print(f"Example: python extract_alnus.py {folder_path} -c {args.city} -a '{available_cols[0]}'")
            else:
                print("No columns found.")
        else:
            min_year = data_df['year'].min()
            max_year = data_df['year'].max()
            print("\nExtracted Data Summary:")
            print(data_df.head())
            print(f"\nTotal records: {len(data_df)}")
            print(f"Year range: {min_year} - {max_year}")
            
            # Create plot
            output_file = f'{allergen_name.lower()}_{args.city.lower()}_{min_year}-{max_year}.png'
            plot_allergen_by_week(data_df, allergen_name, args.years, args.city, output_file)
    else:
        print(f"Invalid folder path: {folder_path}")
