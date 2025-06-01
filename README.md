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
After this the program asks the user to type the directory where to store the file without typing its name it should be like this: C:\Users\Desktop\
After this the program will manage all the data and extract them and store them to the aformentioned directory with name similar to the measurement point itself.
############################################################################################################################################################################################################################################################################################################################################################################################################################################
The second code: Extract Hertta Plant Data: is a python code that extract plant data from Hertta database: the site provides the user with excel sheet that contains the Place,	Date,	Class,	Order,	Family,	Species,	Toxicity,	Calculations,	Size	Kpl/l,	Mass (µg/l),	Mass share%,	Carbon Content (µg/L),	Sea,	Inland Water,	Harmful, Blue Green,	Flagellate,	Nanoplankton.
You have to insert first the directory to where you stored the exported file. 
And then the number of meaurement points, and then their names.
It finally asks the user the directory where to store the file, it should be written like this: C:\Users\Desktop\
the exported files contain the data for each measurement point separately. 


