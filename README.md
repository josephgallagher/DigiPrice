# DigiPrice
Takes a bill of materials in excel and populates the price column. This was originally written to speed up the updating of a whole bunch of purchasing sheets.

This won't be useful in general, but a few adjustments should make this work for an appropriate BOM. This particular scraper assumes that the part numbers are in column #6 of the file, for instance, and places the prices into column #7...

Anyhow, you can use this script to read through a list of Digikey partnumbers, and, using BeatuifulSoup & urllib2, post the data to Digikey and grab the corresponding price table from the html. Then dump the corresponding prices into the corresponding cells.
