#/bin/bash

#nmcli connection edit id preconfigured

read -p "Input IP number 192.168.0.xx: " ipnum

echo ""
echo "IP address is 192.168.0.${ipnum}" 
echo "Press Ctrl-C to cancel."

read Wait

nmcli connection modify preconfigured ipv4.method manual ipv4.addresses 192.168.0.${ipnum}/24
nmcli connection modify preconfigured ipv4.gateway 192.168.0.1
nmcli connection modify preconfigured ipv4.dns 192.168.0.1
