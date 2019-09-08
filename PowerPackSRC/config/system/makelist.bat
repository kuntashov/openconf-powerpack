@echo off
dir /b /oE | perl -n -e "chomp&&/dll|ocx|wsc$/&&print qq(\"$_\",\n)"