#Create a directory for multiple clients on Windows.

import os

List_of_clients = []

path = r'C:\Clients'

error_log = []

for x in List_of_clients:
    try:
        directory_to_make = os.path.join(f'{path}\\{x}')
        os.mkdir(directory_to_make)
    except FileExistsError:
        error_log.append(x)

if len(error_log) > 0:
    print("Following clients already existed:")
    for x in error_log:
        print(x)
