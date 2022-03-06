from csv import excel
from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title='Top Rated Movies' 
print(excel.sheetnames)
sheet.append(['Movie Rank','Movie Name','Year of Release','IMDB Rating'])


try:
    source=requests.get('https://www.imdb.com/chart/top/')
    # This will generate a response object and the response object will be get stored in source
    # source will have the source code of the html web page
    source.raise_for_status() 
    # throws error if url doesnt exist

    soup=BeautifulSoup(source.text,'html.parser') #parses html code stored in source, in the soup variable 
    
    movies=soup.find('tbody',class_="lister-list").find_all('tr')
    
    for movie in movies:

        name = movie.find('td',class_="titleColumn").a.text
        rank=movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        year=movie.find('td',class_="titleColumn").span.text.strip('()')
        rating=movie.find('td',class_="ratingColumn imdbRating").strong.text
        print(rank,name,year,rating)
        sheet.append([rank,name,year,rating])
        



except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')



