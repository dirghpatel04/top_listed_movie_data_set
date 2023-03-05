import pandas as pd
import requests
from bs4 import BeautifulSoup
import openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top IMDB Movie'
# print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name','Year of Relase','IMDB Rating'])
try : 
    url = "https://www.imdb.com/chart/top/"
    r = requests.get(url)
    r.raise_for_status()
    htmlcontent = r.content
    
    soup = BeautifulSoup(htmlcontent , 'html.parser')
    # print(soup.prettify)

    movies = soup.find('tbody',class_ ='lister-list').find_all('tr') # OR text1 = movies.find_all('tr') then print(text)
    # text1 = movies.find_all('tr')
    # print(len(text1))
    
    for movie in movies:
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        name = movie.find('td', class_="titleColumn").a.text
        year = movie.find('td', class_="titleColumn").span.text.strip('()') 
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        # print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])

except Exception as e:
    print(e)
    
excel.save('Top_IMDB_List_of_250_Movie.xlsx')