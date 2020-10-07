# geo_vba

A project to bring geography to VBA. Creating geometric models to make geographical calculations easier.
Building on code examples from [LeafletJS](https://github.com/Leaflet/Leaflet)

## Functions

![Screen 1 - main geographical functions](https://github.com/krijnsent/geo_vba/blob/master/img/screen1.png?raw=true)

The default Excel function for distance is (D8 & D9 holding latitude, E8 and E9 holding longitude):
> =ACOS(COS(RADIANS(90-D9)) *COS(RADIANS(90-D8)) +SIN(RADIANS(90-D9)) *SIN(RADIANS(90-D8)) *COS(RADIANS(E9-E8))) *6370.97327862*1000

The geo_vba version is:
> =geo_distance(D8,E8,D9,E9)

Same goes for surface of a square with 2 known coordinates (NorthEast & SouthWest corner):
> =(6370.97327862*1000*2*PI())^2*ABS((E9-E8)/360*(COS(RADIANS(D9))+COS(RADIANS(D8)))*(D9-D8)/360/2)

The geo_vba version is:
> =geo_surface(D8,E8,D9,E9)

And some more stuff like:
> =geo_earth_radius() (earth radius in kilometers)
> =geo_earth_radius("mi")


![Screen 2 - OSM search & reverse search functions](https://github.com/krijnsent/geo_vba/blob/master/img/screen2.png?raw=true)

Secondly, the formulas can pull in information from OSM nominatim to geocode an address:
> =geo_nom_search(C35)

> =geo_nom_search(C37,,FALSE,"lat,lon,display_name")

> =geo_nom_search(C36,,FALSE,"lat,lon,display_name,osm_type,osm_id")


Or to see what is on a lat&lon location:
> =geo_nom_reverse(C40,D40)

> =geo_nom_reverse(C41,D41,FALSE)

> =geo_nom_reverse(C42,D42,FALSE,"postcode,country_code")


![Screen 3 - extract GPS data from images](https://github.com/krijnsent/geo_vba/blob/master/img/screen3.png?raw=true)

Extract latitude & longitude from an image (if it has that info), and a bit more:
> =geo_img_data(C67)

> =geo_img_data(C69,"DateTime,EquipMake,EquipModel,ExifPixXDim,ExifPixYDim")


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

