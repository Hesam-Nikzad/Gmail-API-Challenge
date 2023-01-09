from Google import Gmail

# ----- Main ----- #

CLIENT_FILE = 'credentials.json'
gmail = Gmail(CLIENT_FILE)
gmail.Send('hessumg@gmail.com', 'Welcome', 'Welcome to the Gmail API challenge.')
gmail.Search('google')