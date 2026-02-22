"""
Flask web application for pollen data visualization with interactive map and plots.
"""

from flask import Flask, render_template, jsonify, request, send_file
from flask_cors import CORS
import json
import pandas as pd
from pathlib import Path
import matplotlib
matplotlib.use('Agg')  # Use non-GUI backend for thread safety
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import mpld3
from io import BytesIO
import base64
import re
import warnings
import pickle
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import signal
import sys

app = Flask(__name__)
CORS(app)

# Configuration
DATA_FOLDER = 'data'
CACHE_FILE = 'cache.pkl'
FRENCH_CITIES_COORDS = {
    'AGEN': {'lat': 44.2049, 'lon': 0.6212},
    'AIXENPRO': {'lat': 43.5297, 'lon': 5.4474},  # Aix-en-Provence
    'AJACCIO': {'lat': 41.9192, 'lon': 8.7386},
    'ALES': {'lat': 44.1249, 'lon': 4.0808},  # Alès
    'AMBERIEU': {'lat': 45.9593, 'lon': 5.3527},  # Ambérieu-en-Bugey
    'AMIENS': {'lat': 49.8941, 'lon': 2.2958},
    'ANDORRA': {'lat': 42.5063, 'lon': 1.5218},  # Andorra la Vella
    'ANGERS': {'lat': 47.4784, 'lon': -0.5632},
    'ANGOULEM': {'lat': 45.6484, 'lon': 0.1562},  # Angoulême
    'ANNECY': {'lat': 45.8992, 'lon': 6.1294},
    'ANNEMASS': {'lat': 46.1944, 'lon': 6.2376},  # Annemasse
    'ANTONY': {'lat': 48.7536, 'lon': 2.2966},
    'AURILLAC': {'lat': 44.9267, 'lon': 2.4417},
    'AUXERRE': {'lat': 47.7982, 'lon': 3.5738},
    'AVIGNON': {'lat': 43.9493, 'lon': 4.8055},
    'BAGNOLS': {'lat': 44.1626, 'lon': 4.6205},  # Bagnols-sur-Cèze
    'BAYONNE': {'lat': 43.4929, 'lon': -1.4748},
    'BESANCON': {'lat': 47.2378, 'lon': 6.0241},
    'BORDEAUX': {'lat': 44.8378, 'lon': -0.5792},
    'BOURGES': {'lat': 47.0810, 'lon': 2.3988},
    'BOURGENB': {'lat': 46.2052, 'lon': 5.2258},  # Bourg-en-Bresse
    'BREST': {'lat': 48.3905, 'lon': -4.4860},
    'BRIANCON': {'lat': 44.8994, 'lon': 6.6422},
    'CAEN': {'lat': 49.1829, 'lon': -0.3707},
    'CASTRES': {'lat': 43.6053, 'lon': 2.2400},
    'CHAMBERY': {'lat': 45.5646, 'lon': 5.9178},
    'CHARLEVI': {'lat': 49.7621, 'lon': 4.7244},  # Charleville-Mézières
    'CHAUMONT': {'lat': 48.1119, 'lon': 5.1396},
    'CHOLET': {'lat': 47.0600, 'lon': -0.8790},
    'CLERMONT': {'lat': 45.7772, 'lon': 3.0862},  # Clermont-Ferrand
    'CORTE': {'lat': 42.3034, 'lon': 9.1496},
    'DIJON': {'lat': 47.3220, 'lon': 5.0398},
    'DINAN': {'lat': 48.4547, 'lon': -2.0486},
    'DOLE': {'lat': 47.0923, 'lon': 5.4890},
    'DRAGUIGN': {'lat': 43.5361, 'lon': 6.4663},  # Draguignan
    'GAP': {'lat': 44.5593, 'lon': 6.0799},
    'GRENOBLE': {'lat': 45.1885, 'lon': 5.7245},
    'LAFLECHE': {'lat': 47.6982, 'lon': -0.0755},  # La Flèche
    'LAROCHE': {'lat': 46.6697, 'lon': -1.4260},  # La Roche-sur-Yon
    'LEMANS': {'lat': 48.0061, 'lon': 0.1996},  # Le Mans
    'LEPUYENV': {'lat': 45.0437, 'lon': 3.8852},  # Le Puy-en-Velay
    'LILLE': {'lat': 50.6292, 'lon': 3.0573},
    'LIMOGES': {'lat': 45.8336, 'lon': 1.2611},
    'LORIENT': {'lat': 47.7483, 'lon': -3.3702},
    'LYON': {'lat': 45.7640, 'lon': 4.8357},
    'LYONII': {'lat': 45.7640, 'lon': 4.8357},
    'LYONIII': {'lat': 45.7640, 'lon': 4.8357},
    'LYONWEST': {'lat': 45.7640, 'lon': 4.8357},
    'MACON': {'lat': 46.3069, 'lon': 4.8287},
    'MARSEILL': {'lat': 43.2965, 'lon': 5.3698},
    'MELUN': {'lat': 48.5392, 'lon': 2.6554},
    'METZ': {'lat': 49.1193, 'lon': 6.1757},
    'MILLAU': {'lat': 44.0997, 'lon': 3.0785},
    'MONTPELL': {'lat': 43.6108, 'lon': 3.8767},  # Montpellier
    'MULHOUSE': {'lat': 47.7508, 'lon': 7.3359},
    'NANCY': {'lat': 48.6921, 'lon': 6.1844},
    'NANTES': {'lat': 47.2184, 'lon': -1.5536},
    'NARBONNE': {'lat': 43.1843, 'lon': 3.0031},
    'NEVERS': {'lat': 46.9896, 'lon': 3.1590},
    'NICE': {'lat': 43.7102, 'lon': 7.2620},
    'NICE2': {'lat': 43.7102, 'lon': 7.2620},
    'NIMES': {'lat': 43.8367, 'lon': 4.3601},
    'NIMES2': {'lat': 43.8367, 'lon': 4.3601},
    'NIORT': {'lat': 46.3237, 'lon': -0.4588},
    'NOUMEA': {'lat': -22.2758, 'lon': 166.4580},
    'ORLEANS': {'lat': 47.9029, 'lon': 1.9093},
    'PARIS': {'lat': 48.8566, 'lon': 2.3522},
    'PARISII': {'lat': 48.8566, 'lon': 2.3522},
    'PAU': {'lat': 43.2951, 'lon': -0.3708},
    'PERPIGNA': {'lat': 42.6887, 'lon': 2.8948},  # Perpignan
    'POITIERS': {'lat': 46.5802, 'lon': 0.3404},
    'REIMS': {'lat': 49.2583, 'lon': 4.0317},
    'RENNES': {'lat': 48.1113, 'lon': -1.6800},
    'ROANNE': {'lat': 46.0362, 'lon': 4.0680},
    'ROUEN': {'lat': 49.4432, 'lon': 1.0993},
    'SAINT-ET': {'lat': 45.4397, 'lon': 4.3872},  # Saint-Étienne
    'SEDAN': {'lat': 49.7019, 'lon': 4.9403},
    'STRASBOU': {'lat': 48.5734, 'lon': 7.7521},
    'TARBES': {'lat': 43.2320, 'lon': 0.0781},
    'TOULON': {'lat': 43.1242, 'lon': 5.9280},
    'TOULOUSE': {'lat': 43.6047, 'lon': 1.4442},
    'TOURS': {'lat': 47.3941, 'lon': 0.6848},
    'TROYES': {'lat': 48.2973, 'lon': 4.0744},
    'TULLE': {'lat': 45.2650, 'lon': 1.7725},
    'VALENCE': {'lat': 44.9334, 'lon': 4.8924},
    'VICHY': {'lat': 46.1286, 'lon': 3.4264}
}

# Global cache
cache = {
    'cities': {},  # {city_name: {'data': df, 'allergens': [list], 'year_range': {min, max}}}
    'year_range': {'min': 2000, 'max': 2026}
}
cache_lock = threading.RLock()  # Thread-safe lock for cache access


def save_cache():
    """Save cache to pickle file."""
    try:
        with open(CACHE_FILE, 'wb') as f:
            pickle.dump(cache, f)
        print(f"✓ Cache saved to {CACHE_FILE}")
    except Exception as e:
        print(f"✗ Error saving cache: {e}")


def load_cache_from_file():
    """Load cache from pickle file if it exists."""
    cache_path = Path(CACHE_FILE)
    if cache_path.exists():
        try:
            with open(CACHE_FILE, 'rb') as f:
                return pickle.load(f)
        except Exception as e:
            print(f"✗ Error loading cache from file: {e}")
    return None


def process_excel_file(excel_file):
    """Process a single Excel file and return cache entry and years."""
    try:
        # Extract city name from filename
        # Format: Particle_Extract_VICHY_2020-01-01_to_2020-12-31
        filename = excel_file.stem.upper()
        
        # Remove "Particle_Extract_" prefix
        city = filename.replace('PARTICLE_EXTRACT_', '')
        
        # Remove date patterns (e.g., "2020-01-01_to_2020-12-31")
        city = re.sub(r'\d{4}-\d{2}-\d{2}.*$', '', city)  # Remove date part onwards
        
        # Clean up
        city = city.strip('_-. ').strip()
        
        if not city:
            print(f"⊘ Skipping {excel_file.name} (empty city name)")
            return None
        
        # Read Excel file - use openpyxl with data_only mode
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
            # data_only=True skips formulas, much faster
            if excel_file.suffix.lower() == '.xls':
                df = pd.read_excel(excel_file, engine='xlrd')
            else:
                df = pd.read_excel(excel_file, engine='openpyxl')
        
        if df.empty or len(df.columns) < 2:
            return None
        
        # Process data
        df_processed = df.copy()
        df_processed.rename(columns={df_processed.columns[0]: 'date'}, inplace=True)
        
        # Parse dates and years
        df_processed['date'] = pd.to_datetime(df_processed['date'], errors='coerce')
        df_processed['year'] = df_processed['date'].dt.year
        df_processed = df_processed.dropna(subset=['date', 'year'])

        if df_processed.empty:
            return None
        
        # Print date and year parsing summary
        print(f"✓ Processed {excel_file.name}: Parsed {len(df_processed)} valid date records, year range {df_processed['year'].min()} - {df_processed['year'].max()}")

        # Extract allergen names from original columns
        allergen_names = [col for col in df.columns[1:] if isinstance(col, str)]
        
        # Pre-compute weekly mean values for each allergen
        allergen_data = {}
        for allergen_name in allergen_names:
            # Calculate weekly means using the actual allergen column name
            df_sorted = df_processed.sort_values('date').copy()
            df_sorted['year_week'] = df_sorted['date'].dt.to_period('W')
            
            # Group by week and calculate mean for this allergen
            if allergen_name in df_sorted.columns:
                weekly_mean = df_sorted.groupby('year_week')[allergen_name].mean()
                # print(f"  - {allergen_name}: Computed weekly means with {len(weekly_mean)} weeks")
                # print(f"    Sample data:\n{weekly_mean.head()}")
                # print timestamps for debugging
                #print(f"    Week timestamps:\n{weekly_mean.index.to_timestamp()}")
                
                # Store weekly means with timestamp index
                if not weekly_mean.empty:
                    allergen_data[allergen_name] = {
                        'weekly_mean': weekly_mean,
                        'week_timestamps': weekly_mean.index.to_timestamp()
                    }
        
        # Return city data and years
        return {
            'city': city,
            'allergen_names': allergen_names,
            'allergen_data': allergen_data,
            'year_range': {
                'min': int(df_processed['year'].min()),
                'max': int(df_processed['year'].max())
            },
            'record_count': len(df_processed),
            'years': set(df_processed['year'].unique())
        }
    
    except Exception as e:
        print(f"✗ Error loading {excel_file.name}: {e}")
        return None

def load_all_data():
    """Load all Excel data into cache or from cache file."""
    # Try loading from cache file first
    cached = load_cache_from_file()
    if cached:
        with cache_lock:
            cache.update(cached)
        print(f"✓ Cache loaded from {CACHE_FILE}")
        return
    
    print("Loading data cache from Excel files (parallel)...")
    data_path = Path(DATA_FOLDER)
    excel_files = list(data_path.glob('*.xls')) + list(data_path.glob('*.xlsx'))
    # keep only first 20 files for faster startup
    #excel_files = excel_files[:20]
    
    years = set()

    # Process files in parallel using ThreadPoolExecutor
    # Increase workers since Excel reading is now optimized
    with ThreadPoolExecutor(max_workers=12) as executor:
        futures = {executor.submit(process_excel_file, excel_file): excel_file 
                   for excel_file in excel_files}
        
        for future in as_completed(futures):
            result = future.result()
            if result:
                city = result['city']
                with cache_lock:
                    if city not in cache['cities']:
                        cache['cities'][city] = {
                            'allergen_names': [],
                            'allergen_data': {}
                        }

                    existing = set(cache['cities'][city]['allergen_names'])
                    for name in result['allergen_names']:
                        if name not in existing:
                            cache['cities'][city]['allergen_names'].append(name)
                            existing.add(name)
                    #cache['cities'][city]['allergen_names'].extend(result['allergen_names'])
                    print(f"✓ Merging allergens for {city}: {len(result['allergen_names'])} allergens, {result['record_count']} records")

                    for allergen_name, data in result['allergen_data'].items():
                        #print(f"  - Processing allergen '{allergen_name}' for {city}")
                        if allergen_name not in cache['cities'][city]['allergen_data']:
                            cache['cities'][city]['allergen_data'][allergen_name] = data
                        else:
                            # merge data if allergen already exists
                            existing_data = cache['cities'][city]['allergen_data'][allergen_name].copy()
                            cache['cities'][city]['allergen_data'][allergen_name]['weekly_mean'] = existing_data['weekly_mean'].combine_first(data['weekly_mean'])
                            cache['cities'][city]['allergen_data'][allergen_name]['week_timestamps'] = existing_data['week_timestamps'].union(data['week_timestamps'])

                    # if year_range is not in cache, initialize it as empty
                    if 'year_range' not in cache['cities'][city]:
                        cache['cities'][city]['year_range'] = {'min': result['year_range']['min'], 'max': result['year_range']['max']}
                        print(f"Initialized year range for {city}: {cache['cities'][city]['year_range']}")

                    if result['year_range']['min'] < cache['cities'][city]['year_range']['min']:
                        cache['year_range']['min'] = result['year_range']['min']
                        cache['cities'][city]['year_range']['min'] = result['year_range']['min']
                        print(f"Updated global min year to {cache['year_range']['min']}")
                    if result['year_range']['max'] > cache['cities'][city]['year_range']['max']:
                        cache['year_range']['max'] = result['year_range']['max']
                        cache['cities'][city]['year_range']['max'] = result['year_range']['max']
                        print(f"Updated global max year to {cache['year_range']['max']}")
                years.update(result['years'])
                print(f"✓ Loaded {city}: {len(result['allergen_names'])} allergens, {result['record_count']} records")
    
    # Update global year range
    with cache_lock:
        if result['year_range']['min'] < cache['year_range']['min']:
            cache['year_range']['min'] = result['year_range']['min']
        if result['year_range']['max'] > cache['year_range']['max']:
            cache['year_range']['max'] = result['year_range']['max']
    
    print(f"Data cache loaded: {len(cache['cities'])} cities")
    print(f"Year range: {cache['year_range']['min']} - {cache['year_range']['max']}")
    
    # Save cache to file
    save_cache()


def get_cities_from_data():
    """Extract all cities from cache."""
    city_data = []
    
    with cache_lock:
        for city_name in cache['cities'].keys():
            if city_name in FRENCH_CITIES_COORDS:
                coord = FRENCH_CITIES_COORDS[city_name]
                city_data.append({
                    'name': city_name,
                    'lat': coord['lat'],
                    'lon': coord['lon']
                })
            else:
                # Try to match partial names
                for key, coord in FRENCH_CITIES_COORDS.items():
                    if city_name[:3] in key or key[:3] in city_name:
                        city_data.append({
                            'name': city_name,
                            'lat': coord['lat'],
                            'lon': coord['lon']
                        })
                        break
    
    return sorted(city_data, key=lambda x: x['name'])


def get_allergens_for_city(city_name):
    """Get list of allergens for a city from cache."""
    city_name = city_name.upper()
    
    with cache_lock:
        if city_name in cache['cities']:
            return cache['cities'][city_name]['allergen_names']
    
    return []


def get_year_range():
    """Get year range from cache."""
    with cache_lock:
        return cache['year_range'].copy()


def extract_allergen_data(city_name, allergen_name):
    """Extract pre-computed weekly mean allergen data for a city from cache."""
    city_name = city_name.upper()
    print(f"Extracting data for city: {city_name}, allergen: {allergen_name}")
    
    with cache_lock:
        if city_name not in cache['cities']:
            return pd.DataFrame()
        print(f"City '{city_name}' found in cache.")
        city_cache = cache['cities'][city_name]
        
        if allergen_name not in city_cache['allergen_data']:
            print(f"Allergen '{allergen_name}' not found. Available: {list(city_cache['allergen_data'].keys())}")
            return pd.DataFrame()
        print(f"Allergen '{allergen_name}' found in cache for city '{city_name}'.")
        allergen_info = city_cache['allergen_data'][allergen_name]
        
        # Create dataframe from pre-computed weekly means
        df_result = pd.DataFrame({
            'date': allergen_info['week_timestamps'],
            'allergen': allergen_info['weekly_mean'].values,
            'year': allergen_info['week_timestamps'].year
        })
    
    return df_result


def generate_plot_image(df, allergen_name, num_years, city_name, min_year=None, max_year=None):
    """Generate a plot and return as base64 encoded image."""
    if df.empty:
        return None
    
    if max_year is None:
        max_year = df['year'].max()
    if min_year is None:
        min_year = max_year - (num_years - 1)

    df_filtered = df[(df['year'] >= min_year) & (df['year'] <= max_year)].copy()
    if df_filtered.empty:
        return None
    
    plt.close('all')
    fig = plt.figure(figsize=(14, 7))
    
    df_sorted = df_filtered.sort_values('date').copy()
    
    # print date range for debugging
    print(f"Plotting data for {city_name} - {allergen_name}:")
    print(f"  - Date range: {df_sorted['date'].min()} to {df_sorted['date'].max()}")
    print(f"  - Number of records: {len(df_sorted)}")
    plt.scatter(df_sorted['date'], df_sorted['allergen'], s=100, alpha=0.6, color='blue', label=f'Mean {allergen_name}')
    plt.plot(df_sorted['date'], df_sorted['allergen'], alpha=0.3, color='blue')
    
    plt.xlabel('Week', fontsize=12, fontweight='bold')
    plt.ylabel(f'{allergen_name} Value', fontsize=12, fontweight='bold')
    plt.title(f'{allergen_name} Values in {city_name} ({min_year}-{max_year})', fontsize=14, fontweight='bold')
    plt.grid(True, alpha=0.3)
    plt.legend()
    
    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%y'))
    
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    
    # Display plot in the html page using mpld3
    #mpld3.save_html(fig,'templates/fig.html')

    # Convert to base64
    img_buffer = BytesIO()
    plt.savefig(img_buffer, format='png', dpi=100, bbox_inches='tight')
    img_buffer.seek(0)
    img_base64 = base64.b64encode(img_buffer.read()).decode()
    plt.close(fig)
    
    return img_base64


# Routes
@app.route('/')
def index():
    """Serve the main page."""
    return render_template('index.html')


@app.route('/api/cities')
def api_cities():
    """Get list of cities from cache."""
    cities = get_cities_from_data()
    return jsonify(cities)


@app.route('/api/allergens/<city>')
def api_allergens(city):
    """Get allergens for a city."""
    allergens = get_allergens_for_city(city)
    return jsonify(allergens)


@app.route('/api/years')
def api_years():
    """Get year range from cache."""
    year_range = get_year_range()
    return jsonify(year_range)


@app.route('/api/plot', methods=['POST'])
def api_plot():
    """Generate and return plot."""
    try:
        data = request.json
        city = data.get('city')
        allergen = data.get('allergen')
        min_year = int(data.get('min_year'))
        max_year = int(data.get('max_year'))
        
        if not city or not allergen:
            return jsonify({'error': 'Missing city or allergen'}), 400
        
        # Extract data
        df = extract_allergen_data(city, allergen)
        #print df head and info for debugging
        # print(f"Extracted data for {city} - {allergen}:")
        # print(df.head())
        # print(df.info())
        
        if df.empty:
            return jsonify({'error': 'No data found'}), 404
        
        # Generate plot
        num_years = max_year - min_year + 1
        img_base64 = generate_plot_image(df, allergen, num_years, city, min_year, max_year)
        
        if img_base64 is None:
            return jsonify({'error': 'Could not generate plot'}), 500
        
        return jsonify({
            'success': True,
            'image': f'data:image/png;base64,{img_base64}'
        })
    
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/api/refresh-cache', methods=['POST'])
def api_refresh_cache():
    """Refresh cache from Excel files."""
    try:
        # Clear existing cache
        with cache_lock:
            cache['cities'].clear()
            cache['year_range'] = {'min': 2000, 'max': 2026}
        
        # Delete cache file
        cache_path = Path(CACHE_FILE)
        if cache_path.exists():
            cache_path.unlink()
            print(f"✓ Cache file deleted")
        
        # Reload from Excel
        load_all_data()
        
        with cache_lock:
            cities_count = len(cache['cities'])
        
        return jsonify({
            'success': True,
            'message': 'Cache refreshed successfully',
            'cities_count': cities_count
        })
    
    except Exception as e:
        print(f"Error refreshing cache: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/debug/cities')
def debug_cities():
    """Inspect cities in cache (debug only)."""
    with cache_lock:
        cities = list(cache['cities'].keys())
    return jsonify({'cities': cities})

@app.route('/debug/cache')
def debug_cache():
    """Inspect cache contents (debug only)."""
    with cache_lock:
        summary = {
            'cities': len(cache['cities']),
            'year_range': cache['year_range'],
            'details': {}
        }
        
        for city, data in cache['cities'].items():
            summary['details'][city] = {
                'allergens': data['allergen_names'],
                'years': data['year_range'],
                'data_points': {allergen: len(info['week_timestamps']) for allergen, info in data['allergen_data'].items()}
#                'full_data': {allergen: {
#                    'week_timestamps': info['week_timestamps'].tolist(),
#                    'weekly_mean': info['weekly_mean'].tolist()
#                } for allergen, info in data['allergen_data'].items()}
            }
    
    return jsonify(summary)

if __name__ == '__main__':
    # Handle graceful shutdown
    def signal_handler(sig, frame):
        print("\n✓ Shutting down gracefully...")
        sys.exit(0)
    
    signal.signal(signal.SIGINT, signal_handler)
    
    # Load all data at startup
    load_all_data()
    app.run(debug=True, port=5000, use_reloader=False)
