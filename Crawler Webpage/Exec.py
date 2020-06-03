import os

itemId = []

# Read URL file
with open("link.txt", "r") as ins:

	for line in ins:
		itemId.append(line.replace('\n', '')) # Store item id into array
ins.close()


itemId = list(filter(None, itemId)) 

for i in range(len(itemId)): # Execute all id
	os.system("Python3 Generator.py " + itemId[i])