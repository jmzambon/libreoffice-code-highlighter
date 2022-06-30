#!/bin/sh

# add libreoffice IDE styles
STYLESREP=codehighlighter/python/pythonpath/pygments/styles
cp customstyles/libreoffice.py "$STYLESREP"

if ! grep -Fq "libreoffice-classic" $STYLESREP/__init__.py; then
    sed -i "/^ *} *$/i \    'libreoffice-classic': 'libreoffice::LibreOfficeStyle'," $STYLESREP/__init__.py
fi

if ! grep -Fq "libreoffice-dark" $STYLESREP/__init__.py; then
    sed -i "/^ *} *$/i \    'libreoffice-dark': 'libreoffice::LibreOfficeDarkStyle'," $STYLESREP/__init__.py
fi


# create oxt file
cd codehighlighter
zip -r codehighlighter2.oxt .
cd ..
mv codehighlighter/codehighlighter2.oxt ./
