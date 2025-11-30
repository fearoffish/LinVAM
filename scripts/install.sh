#!/bin/bash
set -e
. ./configure-uinput-access.sh
pipx install .. --force
sudo cp ../LinVAM.desktop /usr/share/applications/
sudo cp ../linvam/assets/icons/linvam.svg /usr/share/icons/hicolor/scalable/apps/
