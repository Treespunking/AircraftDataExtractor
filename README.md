This Python script is an **Aircraft Data Extraction System** that uses AI to automatically extract structured information from aircraft listing descriptions. Here's a comprehensive breakdown:

## Core Purpose
The script processes Excel files containing aircraft listing descriptions and extracts specific aviation-related data points using an AI model, then outputs the results to a CSV file with calculated maintenance metrics.

## Key Components

### 1. **AircraftDataExtractor Class**
The main class that handles the entire extraction workflow:

- **Initialization**: Sets up API authentication headers for OpenRouter AI service
- **Data Processing Pipeline**: Takes raw text descriptions and returns structured JSON data with calculated fields

### 2. **AI Integration (OpenRouter)**
- Uses the `qwen/qwen3-235b-a22b:free` model through OpenRouter's API
- Sends carefully crafted prompts to extract specific aircraft data fields
- Implements robust error handling for API failures

### 3. **Data Fields Extracted**
The script extracts 14 specific aviation data points:
- **Basic Info**: Date posted, manufacture year, registration number
- **Engine Metrics**: TTAF (Total Time Airframe), TSN (Time Since New), CSN (Cycles Since New)
- **Maintenance Data**: Time since overhaul, HSI inspection details, maintenance programs
- **Dates**: Last overhaul, overhaul due dates, HSI dates

### 4. **Advanced Calculations**
The script performs sophisticated aviation maintenance calculations:

- **Time Remaining Before Overhaul**: Calculates based on multiple factors:
  - Insurance maintenance programs (8000 hours)
  - HSI-based calculations (4000 hour midlife)
  - Time since new calculations
  - Condition-based assessments

- **Operational Projections**: 
  - Years left for operation
  - Average hours remaining (based on 450 hours/year usage)
  - On-condition repair status determination

### 5. **Robust JSON Parsing**
Implements a three-tier parsing strategy:
1. Standard JSON parsing
2. Fallback decoder for partial JSON
3. Truncation-based recovery for incomplete responses

### 6. **File Processing Workflow**
- Reads Excel input file (`Test42Inputs.xlsx`)
- Processes each description row individually
- Outputs to CSV with predefined field structure
- Handles missing or empty descriptions gracefully

## Technical Features

### Error Handling
- API request failures
- JSON parsing errors
- Data type conversion issues
- Missing environment variables

### Data Safety
- Safe integer conversion with fallback to None
- Date parsing with error handling
- Key renaming to avoid CSV compatibility issues

### Configuration Management
- Uses environment variables for API keys
- Configurable model selection
- Customizable output formatting

## Use Cases
This script would be valuable for:
- **Aircraft Brokers**: Standardizing listing data
- **Maintenance Companies**: Assessing aircraft condition
- **Fleet Managers**: Evaluating acquisition opportunities
- **Insurance Companies**: Risk assessment based on maintenance history

## Dependencies
- `requests`: API communication
- `pandas`: Excel file processing
- `python-dotenv`: Environment variable management
- Standard libraries: `json`, `csv`, `datetime`, `os`
