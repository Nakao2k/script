#/bin/bash

read -p "Input hostname: " name

sudo hostnamectl set-hostname ${name}

exit
