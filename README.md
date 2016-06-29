# dbSNO_scraper
A program to scrape the DataBase of Cysteine S-NitrOsylation and add experiment type to a local excel file. Great for automating the tedious process of manually searching each protein and figuring out what to add. Can be modified to obtain S-nitrosylated Peptide, Secondary Structure of S-nitrosylated Peptide, Solvent Accessibility of nitrosylated Site, Substrate Motifs, PubMed ID, and Experiment type. Currently configured for Experiment type.

To start, modify fetch_DBSNO.py to change the workbook input, output, and sheet name. Modify the row and column variables to change it to your directory that you want to lookup. After that, run fetch_DBSNO.py and it will query dbSNO v2.0 and download all of the .HTML files locally.

To retreive relevant data of experiment type, modify parse_DBSNO.py to match your local excel file. It will then search the database of local .HTML files and add the experiment type in whatever column specified.

Hamaad Markhiani
hamaad.markhiani@gmail.com
