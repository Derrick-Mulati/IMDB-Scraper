from bs4 import BeautifulSoup
import requests, openpyxl

# Create Excel file
excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])

# IMDb URL
url = 'https://www.imdb.com/chart/top/'
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

try:
    # Request data
    source = requests.get(url, headers=headers, timeout=10)
    source.raise_for_status()
    soup = BeautifulSoup(source.text, 'html.parser')

    # Extract movies
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])

    # Save the file
    excel.save('IMDB Movie Ratings.xlsx')
    print("Data saved successfully!")

except requests.exceptions.RequestException as e:
    print("Error fetching data:", e)
