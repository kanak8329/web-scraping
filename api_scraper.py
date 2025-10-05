import requests
import pandas as pd
import time


def get_all_universities():
    """
    Hits the main QS Rankings API to get the core data for all ranked universities.
    """
    # This URL is the source of data for the main rankings table.
    # It fetches up to 2500 ranked institutions.
    api_url = "https://www.topuniversities.com/api/qs-rankings/en/2025/916481?qs_ranking_instance_id=916481&items_per_page=2500&page=0"

    print("➡️ Accessing the main university database API...")

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
    }

    response = requests.get(api_url, headers=headers)

    if response.status_code != 200:
        print(f"❌ Failed to fetch the main university list. Status code: {response.status_code}")
        return []

    data = response.json()
    universities = data.get('data', [])

    print(f"✅ Found {len(universities)} universities in the database.")
    return universities


def get_detailed_stats(university_id):
    """
    Hits the specific API for one university to get its detailed student/staff stats.
    """
    api_url = f"https://www.topuniversities.com/api/institution/en/{university_id}"

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'
    }

    try:
        response = requests.get(api_url, headers=headers, timeout=10)
        if response.status_code == 200:
            return response.json()
    except requests.RequestException:
        return None
    return None


def parse_stats(details):
    """
    Parses the JSON response from the detailed stats API to find the specific data points required.
    """
    parsed_data = {}

    stats_map = {
        'stats_total_student': 'Total Students (Total)',
        'stats_ug_student': 'Total Students (UG students)',
        'stats_pg_student': 'Total Students (PG students)',
        'stats_total_inter_student': 'International Students (Total)',
        'stats_ug_inter_student': 'International Students (UG students)',
        'stats_pg_inter_student': 'International Students (PG students)',
        'stats_total_faculty': 'Total Faculty Staff (Total)',
        'stats_dom_faculty': 'Total Faculty Staff (Domestic staff)',
        'stats_int_faculty': 'Total Faculty Staff (Int\'l staff)',
    }

    if 'stats' in details:
        for stat in details['stats']:
            api_key = stat.get('type')
            if api_key in stats_map:
                display_key = stats_map[api_key]
                parsed_data[display_key] = stat.get('value')

    return parsed_data


def main():
    """
    Main function to run the scraper.
    """
    universities = get_all_universities()

    if not universities:
        print("Could not retrieve university list. Exiting.")
        return

    all_data = []
    total_universities = len(universities)

    for i, uni in enumerate(universities):
        university_id = uni.get('nid')
        university_name = uni.get('title')

        print(f"⚙️ Processing ({i + 1}/{total_universities}): {university_name}")

        details = get_detailed_stats(university_id)

        final_record = {'University Name': university_name}

        if details:
            parsed_stats = parse_stats(details)
            final_record.update(parsed_stats)

        all_data.append(final_record)
        time.sleep(0.1)  # Be respectful with a small delay between requests

    print("\n" + "=" * 50)
    print("💾 Scraping complete. Saving data to Excel file...")

    df = pd.DataFrame(all_data)

    # Define the desired column order
    cols_order = [
        'University Name', 'Total Students (Total)', 'Total Students (UG students)', 'Total Students (PG students)',
        'International Students (Total)', 'International Students (UG students)',
        'International Students (PG students)',
        'Total Faculty Staff (Total)', 'Total Faculty Staff (Domestic staff)', 'Total Faculty Staff (Int\'l staff)'
    ]
    df = df.reindex(columns=cols_order)

    filename = 'university_api_data.xlsx'
    df.to_excel(filename, index=False)

    print(f"🎉 SUCCESS! Data for {len(all_data)} universities saved to '{filename}'")


if __name__ == "__main__":
    main()