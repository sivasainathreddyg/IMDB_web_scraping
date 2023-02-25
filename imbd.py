import requests,openpyxl
from bs4 import BeautifulSoup

excel=openpyxl.Workbook()
print(excel.sheetnames)
Sheet=excel.active
Sheet.tille='Top Rated Movies'
print(excel.sheetnames)
Sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB Rating'])



try:
    source=requests.get("https://www.imdb.com/chart/top/")
    source.raise_for_status()

    soup=BeautifulSoup(source.text,'html.parser')
    
    movies=soup.find('tbody',class_="lister-list").find_all('tr')
    print(len(movies))
    for movie in movies:
        name=movie.find('td',class_='titleColumn').a.text
        rank=movie.find('td',class_='titleColumn').get_text(strip=True).split('.')[0]
        year=movie.find('td',class_='titleColumn').span.text.strip('()')
        rating=movie.find('td',class_='ratingColumn imdbRating').strong.text 
        print(rank,name,year,rating)
        Sheet.append([rank,name,year,rating])
        


except Exception as e:
    print(e)

excel.save(' Top 250 IMBD movies.xlsx')