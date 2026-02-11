# Popular-Movies-Scrapper

A small Python script that scrapes Rotten Tomatoes’ *Popular Movies* editorial page and exports the current list of popular streaming movies into an Excel file.

This project was created as a **quick refresher on web scraping in Python**, not as a production-ready scraper. Putting it on GitHub just for fun.

## What It Does

- Scrapes the current list of popular movies from Rotten Tomatoes
- Extracts:
  - Movie title  
  - Rotten Tomatoes score  
  - Link to the movie’s Rotten Tomatoes page  
  - Poster image link  
- Saves the data to `PopularMoviesToday.xlsx`
- Applies basic formatting to the Excel file:
  - Auto-sized columns
  - Percentage formatting for scores
  - Clickable hyperlinks
  - Styled table layout

## Notes

- The scraper relies on Rotten Tomatoes’ current page structure and may break if the site layout changes.
- Movies without a listed score are assigned a value of `0%`. (I know, that's not what you're supposed to do. I'll deal with it another time)
