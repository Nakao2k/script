#/bin/bash

read -p "Input IP number 192.168.0.xx: " ipnum

echo ""
echo "IP address is 192.168.0.${ipnum}" 
echo "Press Ctrl-C to cancel."

read Wait

nmcli connection add type ethernet con-name ethstatic
nmcli connection modify ethstatic ipv4.method manual ipv4.addresses 192.168.0.${ipnum}/24
nmcli connection modify ethstatic ipv6.method manual ipv6.addresses 2001:dead:1::${ipnum}

nmcli connection modify ethstatic ipv4.gateway 192.168.0.1/24
nmcli connection modify ethstatic ipv6.gateway 2001:dead:1::beef

nmcli connection up ethstatic ifname ens192

sudo hostnamectl set-hostname redhat-${ipnum}

echo "ip and hostname changed."
hostname

