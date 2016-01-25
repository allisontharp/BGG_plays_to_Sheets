import gspread
import json
from oauth2client.client import SignedJwtAssertionCredentials
from boardgamegeek import BoardGameGeek
from datetime import datetime
from functs import write_cell_colstring, col

#-- Attention!  Be sure to edit these values! --#
json_file = 'api.json' # your json file name (name.json)
worksheet = "boardgame_test" # worksheet name
spreadsheet = "Sheet1" # spreadsheet name that you want the data to go into
bgg_username = "mad4hatter" # your bgg username
username = "Allison"    # your name
oldest_game = datetime(2015,12,31,0,0,0) # year, month, day of the oldest game you want to import (only valid for first time running the script)

# connect to spreadsheet
json_key = json.load(open(json_file))
scope = ['https://spreadsheets.google.com/feeds']
credentials = SignedJwtAssertionCredentials( json_key['client_email'], json_key['private_key'], scope)

gc = gspread.authorize(credentials)

sh = gc.open(worksheet)
sheet1 = sh.worksheet("Sheet1")


# connect to BGG
bgg = BoardGameGeek()
userplays = bgg.plays(bgg_username).plays

# most recent game in spreadsheet
gamedates = sheet1.col_values(1)
oldest_game.strftime('%m/%d/%Y')

if gamedates[len(gamedates)-1] == 'date':
    latestgame = oldest_game
else:
    latestgame = gamedates[len(gamedates)-1]
    latestgame = datetime.strptime(latestgame,'%m/%d/%Y')

lastrow = len(gamedates)

headers = sheet1.row_values(1)

oldest_game = oldest_game.strftime('%m/%d/%Y')


for games in userplays:
        
    if games.date > latestgame: 

            cell_values = ['']*len(headers)
            temp = games.players
            
    
            lastrow += 1
                
            cell_values[col('date',headers)] = games.date.strftime('%m/%d/%Y')
            cell_values[col('gamename',headers)] = games.game_name
            cell_values[col('duration',headers)] = str(games.duration)
            
                
            print games.game_name, games.duration
                
            winners = []
            win = '1'
            [winners.append(players.name) for players in games.players if players.win == win]
            winnersfinal = ", ".join(winners)
                
            cell_values[col('winners',headers)] = winnersfinal
            
            useridx = [item for item in range(len(games.players)) if games.players[item].name == username]                
    
            newplayers = ''
            if games.players[useridx[0]].new == '1':
                newplayers += 'New'
                    
            cell_values[col('comment',headers)] = newplayers
                
            if winners != []:
                winscore = [float(players.score.replace(" ","")) for players in games.players if players.name == winners[0]]
                        
                
            for players in games.players: 
                try:
                    headers.index(players.name.replace(' ','').lower())
                except:
                    sheet1.update_cell(1,len(headers)+1,players.name.replace(' ','').lower())
                    cell_values.append('')
    
                headers = sheet1.row_values(1)
                cell_values[col(players.name.replace(' ','').lower(), headers)] = players.score
                
            headers = sheet1.row_values(1)
            lastcol = len(headers)
            if lastcol < 26:
                lastcol = chr(lastcol%26 -1 + ord('a'))
            else:
                lastcol = chr(lastcol/26 - 1 + ord('a')) + chr(lastcol%26 -1 + ord('a'))
                
            
            cell_list = sheet1.range('a' + str(lastrow) + ':' + lastcol + str(lastrow))
                
            for i, val in enumerate(cell_values):
                    cell_list[i].value = val
                
            sheet1.update_cells(cell_list)
            print "Game Added"
           
    
    elif games.date == latestgame:
        print "games on latest date need to be added"
    else:
        break
            
            
            
            
            
