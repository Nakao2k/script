#/bin/bash

read -p "Input IP number 192.168.0.xx: " ipnum

echo ""
echo "IP address is 192.168.0.${ipnum}" 
echo "Press Ctrl-C to cancel."

read Wait

sudo sed -i -r "s/^(\s{6}-\s192\.168\.0\.)[0-9]+(\/24)/\1${ipnum}\2/" /etc/netplan/11-installer-config.yaml
sudo hostnamectl set-hostname ubuntu-${ipnum}

echo "ip and hostname changed."
hostname
sudo netplan apply
