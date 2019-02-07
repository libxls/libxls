#!/bin/bash

if [[ $TRAVIS_OS_NAME == 'osx' ]]; then
    brew link --force gettext
fi
autoreconf -i -f
