#! /bin/sh

NAME=WatchingWindow
VERSION=`cat "VERSION"`

zip -9 -o $NAME-$VERSION.oxt \
  META-INF/* \
  description.xml \
  descriptions/* \
  icons/* dialogs/* \
  registration.components \
  pythonpath/**/*.py pythonpath/**/**/*.py \
  *.xcu *.xcs registration.py \
  resources/* \
  help/**/* help/**/**/* \
  watchingwindow.py \
  LICENSE CHANGES NOTICE Translators.txt

