import requests
import json
from dotenv import load_dotenv
import os
from datetime import datetime
import csv
import pandas as pd  # For reading Excel files

class AircraftDataExtractor:
    def __init__(self, api_key):
        self.api_key = api_key
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://your-site.com",      # Optional
            "X-Title": "Your Site Name",                # Optional
        }

    def extract_data(self, text):
        prompt = self._build_prompt(text)
        raw_response = self._call_openrouter(prompt)
        extracted_json = self._parse_response(raw_response)
        if not extracted_json:
            return {}
        calculated_data = self._calculate_fields(extracted_json)
        return calculated_data

    def _build_prompt(self, text):
        return f"""
You are an aircraft data extraction assistant. Given this aircraft listing:
{text}
Please extract the following fields in JSON format using ONLY these exact keys:
"Date advertisement was posted",
"Manufacture Year of plane",
"Registration number of plane",
"TTAF",
"Position of engine",
"TSN",
"CSN",
"Total Time Since Overhaul (TSOH)",
"Time Before Overhaul provided in the information (Early TBO)",
"Hours since HSI (Hot Service Inspection)",
"Date of Last HSI (Hot Service Inspection)",
"Insurance Maintenance Program the engine is enrolled in",
"Date of Last Overhaul",
"Date of Overhaul Due",
Do NOT use any extra description or formatting. Return only valid JSON.
"""

    def _call_openrouter(self, prompt):
        payload = {
            "model": "qwen/qwen3-235b-a22b:free",
            "messages": [
                {"role": "user", "content": prompt}
            ],
            "temperature": 0
        }
        try:
            response = requests.post(
                url="https://openrouter.ai/api/v1/chat/completions", 
                headers=self.headers,
                data=json.dumps(payload)
            )
            if response.status_code == 200:
                return response.json()["choices"][0]["message"]["content"]
            else:
                print(f"API Error: {response.status_code}")
                print(response.text)
                return None
        except Exception as e:
            print(f"Request failed: {e}")
            return None

    def _parse_response(self, raw_text):
        if not raw_text:
            print("Empty response from model.")
            return {}
        raw_text = raw_text.strip()
        try:
            return json.loads(raw_text)
        except json.JSONDecodeError:
            print("Standard JSON parsing failed. Trying fallback parser...")
        decoder = json.JSONDecoder()
        try:
            result, index = decoder.raw_decode(raw_text)
            print("Warning: Partial JSON parsed successfully up to position", index)
            return result
        except Exception as e:
            print("Failed to parse JSON even with fallback decoder:", e)
        last_brace = raw_text.rfind("}")
        if last_brace != -1:
            potential_json = raw_text[:last_brace + 1]
            try:
                result = json.loads(potential_json)
                print("Warning: Successfully parsed by truncating incomplete JSON")
                return result
            except json.JSONDecodeError:
                pass
        print("Failed to parse JSON. Raw response:")
        print(raw_text)
        return {}

    def _safe_int(self, val):
        try:
            return int(val)
        except (ValueError, TypeError):
            return None

    def _safe_date(self, val):
        try:
            return datetime.strptime(val.strip(), "%Y-%m-%d")
        except Exception:
            return None

    def _calculate_fields(self, data):
        tsn = self._safe_int(data.get("TSN"))
        tsOh = self._safe_int(data.get("Total Time Since Overhaul (TSOH)"))
        hsi_hours = self._safe_int(data.get("Hours since HSI (Hot Service Inspection)"))
        engine_program = data.get("Insurance Maintenance Program the engine is enrolled in")

        basis_of_calculation = None
        time_remaining_before_overhaul = None

        if engine_program:
            time_remaining_before_overhaul = 8000
            basis_of_calculation = "Insurance Maintenance Program"
        elif hsi_hours is not None:
            time_remaining_before_overhaul = max(0, 4000 - hsi_hours)
            basis_of_calculation = "Midlife Calculation"
        elif tsOh is not None:
            time_remaining_before_overhaul = max(0, 4000 - tsOh)
            basis_of_calculation = "Midlife Calculation"
        elif tsn is not None:
            if tsn < 8000:
                time_remaining_before_overhaul = 8000 - tsn
                basis_of_calculation = "time since new"
            else:
                time_remaining_before_overhaul = 0
                basis_of_calculation = "condition based"

        date_ad_posted = self._safe_date(data.get("Date advertisement was posted"))
        date_last_overhaul = self._safe_date(data.get("Date of Last Overhaul"))
        date_overhaul_due = self._safe_date(data.get("Date of Overhaul Due"))

        years_left_for_operation = None
        if date_overhaul_due and date_ad_posted:
            years_left_for_operation = round((date_overhaul_due - date_ad_posted).days / 365, 2)
        elif date_last_overhaul and date_ad_posted:
            years_left_for_operation = round((date_ad_posted - date_last_overhaul).days / 365, 2)

        avg_hours_left = None
        if years_left_for_operation is not None:
            avg_hours_left = round(years_left_for_operation * 450, 2)

        on_condition_repair = False
        if (
            tsn is not None and tsn > 8000 and
            tsOh is None and
            hsi_hours is None and
            data.get("Date of Last HSI") is None and
            date_last_overhaul is None
        ):
            on_condition_repair = True

        # Rename problematic keys
        if "Insurance Maintenance Program the engine is enrolled in" in data:
            data["Engine Maintenance Insurance Program Name"] = data.pop("Insurance Maintenance Program the engine is enrolled in")
        if "Date of Last HSI" in data:
            data["Date of Last HSI (Hot Service Inspection)"] = data.pop("Date of Last HSI")
        if "Hours since HSI" in data:
            data["Hours since HSI (Hot Service Inspection)"] = data.pop("Hours since HSI")

        data["Time Remaining before Overhaul"] = time_remaining_before_overhaul
        data["Basis of Calculation"] = basis_of_calculation
        data["years left for operation"] = years_left_for_operation
        data["Average Hours left for operation according to 450 hours annual usage"] = avg_hours_left
        data["On Condition Repair"] = on_condition_repair

        return data


# ————— MAIN EXECUTION —————
if __name__ == "__main__":
    load_dotenv()
    openrouter_api_key = os.getenv("OPENROUTER_API_KEY")

    if not openrouter_api_key:
        raise ValueError("OPENROUTER_API_KEY not found in environment variables.")

    extractor = AircraftDataExtractor(openrouter_api_key)

    input_file = "Test42Inputs.xlsx"
    output_file = "aircraft_output.csv"

    df = pd.read_excel(input_file)

    fieldnames = [
        "Date advertisement was posted",
        "Manufacture Year of plane",
        "Registration number of plane",
        "TTAF",
        "Position of engine",
        "TSN",
        "CSN",
        "Total Time Since Overhaul (TSOH)",
        "Time Before Overhaul provided in the information (Early TBO)",
        "Hours since HSI (Hot Service Inspection)",
        "Date of Last HSI (Hot Service Inspection)",
        "Engine Maintenance Insurance Program Name",
        "Date of Overhaul Due",
        "Date of Last Overhaul",
        "Time Remaining before Overhaul",
        "Basis of Calculation",
        "years left for operation",
        "Average Hours left for operation according to 450 hours annual usage",
        "On Condition Repair"
    ]

    with open(output_file, mode='w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for idx, row in df.iterrows():
            description = str(row.get("Description", ""))
            if not description.strip():
                continue
            print(f"\nProcessing Description {idx + 1}...")
            result = extractor.extract_data(description)

            # Ensure only known keys are written
            cleaned_result = {key: result.get(key) for key in fieldnames}
            writer.writerow(cleaned_result)

    print(f"\nData successfully written to '{output_file}'")