from bs4 import BeautifulSoup
import requests, openpyxl


excel= openpyxl.Workbook()   #to create an excel file where we can store the data that we scraped 

sheet = excel.active
sheet.title= 'Top rated movies'
print(excel.sheetnames)

sheet.append(['Movie Rank','Movie Name','Year of release','Rating of the movie'])



try:                            # try and except for handeling any exception/ error that migh occur
    source = requests.get('https://www.imdb.com/chart/top/')    # to read the HTML source code and scrape the data
    source.raise_for_status()    #shows exception when an invalid url is provided
    soup= BeautifulSoup(source.text,'html.parser')                #to parse the html code that we have retrieved
    
    movies = soup.find('tbody', class_ ="lister-list").find_all('tr')       #to choose that particular part of the page that has the list and rating of the movies i.e from 'tbody' tag we have entered the 'tr' tag which directly has the name of the movies and there are 250 such 'tr' tags i.e. 1 'tr' tag for each movie
    
    for movie in movies:
        name= movie.find('td', class_="titleColumn").a.text    #we enter the 'td' tag here and then the 'a' tag and then the text of the 'a' tag
        
        rank= movie.find('td', class_="titleColumn").get_text(strip= True).split('.')[0]
        
        year= movie.find('td', class_="titleColumn").span.text.strip('()')

        rating= movie.find('td', class_="ratingColumn imdbRating").strong.text
        
        print(rank, name, year, rating)

        sheet.append([rank, name, year, rating])




except Exception as e:
    print(e)

    


excel.save('IMDB movie ratings.xlsx')