# geo_vba

A project to bring geography to VBA. Creating geometric models to make geographical calculations easier.
Building on code examples from [LeafletJS](https://github.com/Leaflet/Leaflet)

## Functions
- Added some excel functions to e.g. calculate distance between two points and the surface of a LatLng Box.
- OpenStreetMap functions, using nominatim for search of an address or reverse-search of a lat/lon combination


# TODO General direction of code:

## Create a scatterplot with a map background
Input lat&lon min&max
- array of tiles -> almost there :-)
- download all tiles
- merge all tiles
- crop image
- set as chart background

## OSM
Add OpenStreetMap functions:
- using overpass-turbo
- getting tiles

## License
[GNU General Public License v3.0](https://choosealicense.com/licenses/gpl-3.0/)

## Inspired by
- http://www.cpearson.com/excel/LatLong.aspx
- https://www.atsunday.com/2012/03/calculating-land-lot-area-with-gps.html

