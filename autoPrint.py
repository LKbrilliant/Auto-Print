import win32com.client as win32 
import os
import time

currentPath = os.getcwd()

print("     _         _          ____       _       _   ")
print("    / \  _   _| |_ ___   |  _ \ _ __(_)_ __ | |_ ")
print("   / _ \| | | | __/ _ \  | |_) | '__| | '_ \| __|")
print("  / ___ \ |_| | || (_) | |  __/| |  | | | | | |_ ")
print(" /_/   \_\__,_|\__\___/  |_|   |_|  |_|_| |_|\__| @LKBrilliant")
print("Print multiple coppies of a one page MS Word document with a unique ID on each page.")
print("The Word document must have a 'bookmark' named 'ID' where the unique ID need to be placed.")
print("Compatible MS Word format: .doc")
print("Quit: Ctrl+c\n")

while True:
    try:
        docName = input("Name of the Document: ")
        if not os.path.isfile("{}.doc".format(docName)):
            print("{}.doc is not in this directory".format(docName))
        else: break
    except KeyboardInterrupt:
        print("\nQuitting...")
        exit()
while True:
    try:
        numPrints = int(input("Number of prints: "))
        break
    except ValueError:
        print("Enter a valid answer!")
    except KeyboardInterrupt:
        print("\nQuitting...")
        exit()
while True:
    try:
        ID = int(input("Starting ID: "))
        break
    except ValueError:
        print("Enter a valid answer!")
    except KeyboardInterrupt:
        print("\nQuitting...")
        exit()

print("Sending documents to print...")
while numPrints >= 1:
    wordApp = win32.gencache.EnsureDispatch('Word.Application')
    wordApp.Visible = False
    doc = wordApp.Documents.Open("{}\\F002.doc".format(currentPath))
    rng = doc.Bookmarks("ID").Range             # An bookmark is placed where the changes need to happen
    rng.InsertAfter("{:02d}".format(ID))
    doc.PrintOut()
    rng.Delete()
    doc.Save()                                  # Save the document, otherwise MS Word will prompt a message to save
    wordApp.Quit()
    time.sleep(3)
    numPrints -= 1
    ID += 1 
print("Done")

