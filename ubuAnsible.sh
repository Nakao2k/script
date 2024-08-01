#!/bin/bash

source /etc/os-release

if [ ${NAME,,} != "ubuntu" ]; then
    echo "This script is for ubuntu."
    exit 1
fi

apt update
apt install software-properties-common
add-apt-repository --yes --update ppa:ansible/ansible
apt install ansible

