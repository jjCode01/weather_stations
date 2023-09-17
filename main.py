import re
from datetime import date, datetime

import requests
from openpyxl import Workbook
from tqdm import tqdm

from src.countrys import noaa_countrys
from src.us_states import us_states

URL = r"http://www.ncdc.noaa.gov/homr/services/station/search"


def get_weather_station(**params) -> list[dict]:
    response = requests.get(URL, params=params)

    if response.status_code != 200:
        return []

    return response.json()["stationCollection"]["stations"]


def get_station_data(station: dict, country: str, state: str) -> list:
    stn_id = station["ncdcStnId"]
    stn_name = station.get("header", {}).get("preferredName", "")
    stn_lat = station.get("header", {}).get("latitude_dec", "")
    stn_lon = station.get("header", {}).get("longitude_dec", "")
    stn_prec = station.get("header", {}).get("precision", "")
    stn_start = station.get("header", {}).get("por", {}).get("beginDate", "")
    stn_end = station.get("header", {}).get("por", {}).get("endDate", "")
    return [
        stn_id,
        country,
        state,
        stn_name,
        stn_lat,
        stn_lon,
        stn_prec,
        to_datetime(stn_start),
        to_datetime(stn_end),
    ]


def to_datetime(val: str) -> date | str:
    if val == "Present":
        return val
    dt_fmt = re.compile(r"^\d{4}-\d{2}-\d{2}(?=T)")
    if dt := dt_fmt.search(val):
        return datetime.strptime(dt.group(), "%Y-%m-%d").date()
    return "Unknown"


def all_stations_to_excel():
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(  # type: ignore
        [
            "StnID",
            "Country",
            "State",
            "Name",
            "Latitue",
            "Longitude",
            "Precision",
            "Begin Date",
            "End Date",
        ]
    )

    print("Compiling US Weather Station Data...")
    usa = tqdm(us_states)
    for state in usa:
        usa.set_description(state)
        params = {"headersOnly": "true", "state": state}
        stations = get_weather_station(**params)

        for stn in stations:
            sheet.append(get_station_data(stn, "USA", state)) # type: ignore

    print("Compiling World Weather Station Data...")
    world = tqdm(noaa_countrys)
    for country in world:
        world.set_description(country)
        params = {"headersOnly": "true", "country": country}
        stations = get_weather_station(**params)

        for stn in stations:
            sheet.append(get_station_data(stn, country, "")) # type: ignore

    dt = datetime.now().strftime("%Y-%m-%d-%H%M%S")

    workbook.save(f"noaa_weather_stations_{dt}.xlsx")


if __name__ == "__main__":
    all_stations_to_excel()
