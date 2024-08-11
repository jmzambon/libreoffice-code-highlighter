#!/bin/bash

# generate-mo-files.sh
for PO_FILE in locales/*/LC_MESSAGES/*.po
do
    MO_FILE="${PO_FILE/.po/.mo}"
    msgfmt -o "codehighlighter/$MO_FILE" "$PO_FILE"
done

# generate codehighlighter2.oxt
cd codehighlighter

zip -r codehighlighter2.oxt .
cd ..
mv codehighlighter/codehighlighter2.oxt ./
