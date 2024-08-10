#!/bin/bash

cd codehighlighter

# generate-mo-files.sh
for PO_FILE in locales/*/LC_MESSAGES/*.po
do
    echo $PO_FILE
    MO_FILE="${PO_FILE/.po/.mo}"
    msgfmt -o "$MO_FILE" "$PO_FILE"
done

# generate codehighlighter2.oxt
zip -r codehighlighter2.oxt .
cd ..
mv codehighlighter/codehighlighter2.oxt ./
