Python file's name: Extract Hertta Final Right, enables us to access the water quality data that are organized on Excel sheet after selecting the needed chemicals like:
Oxygen (Happi liukonen) or Total Nitrogen(Kokonaistyyppenä)...etc
The file in general is organized with column's head as follows: 
column = ['Paikan nimi', 'Paikan ID-numero', 'Paikan ETRS-koord itä', 'Paikan ETRS-koord pohj',	'Paikan syvyys (m)', 'Näytteenottoaika', 'Näytesyvyys',	'Määrityskoodi', 'Suure(suomeksi)', 'Lippu', 'Alkuperäinen arvo', 'Yksikkö', 'Tulos']
The needed columns in general can be: Paikan nimi: place's name; Näytesyvyys: Sample's depth; Suure(suomeksi): Quantity; Näytteenottoaika: Sample's Time; Yksikkö: liek ug/L or mg/L; and Tulos: which is the mass.
The code asks the user first to type the location of the file and where it exists like: C:USERES....xlsx
And then asks the number of sheets inside the file which is generally one
Then asks the name of the sheet
The code asks the user about the column's range which is in general from A:M unless the user asked additional data from Hertta.
Then asks the number of nutrients or chemicals the users like to insert
Then asks the user the name of each or these nutrients the user should provide the names as they are in the excel sheet so using the Finnish language like: ph, Lämpötila...etc
Then he asks the user to insert the number of measurement places in the lake
Then to type their names for instance in my project there were measurements points like: Kallavesi 25, Kallavesi 375...etc
After this the program asks the user to type the direcory where to store the file without typing its name it should be like this: C:\Users\Desktop\
After this the program will manage all the data and extract them and store them to the aformentioned directory with name similar to the measurement point itself. 
