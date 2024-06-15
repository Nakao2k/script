#/bin/bash

read -p "Input IP number 192.168.0.xx: " ipnum

echo ""
echo "IP address is 192.168.0.${ipnum}" 
echo "Press Ctrl-C to cancel."

read Wait

nmcli connection add type ethernet con-name ethMusen
nmcli connection modify ethMusen ipv4.method manual ipv4.addresses 192.168.0.${ipnum}/24
nmcli connection modify ethMusen ipv4.gateway 192.168.0.1
nmcli connection modify ethMusen ipv4.dns 192.168.0.1

echo "ip address changed."

nmcli connection up ethMusen ifname wlan0

