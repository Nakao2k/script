#/bin/bash

read -p "Input IP number 192.168.0.xx: " ipnum

echo ""
echo "IP address is 192.168.0.${ipnum}" 
echo "Press Ctrl-C to cancel."

read Wait

nmcli connection add type ethernet con-name ethStatic
nmcli connection modify ethStatic ipv4.method manual ipv4.addresses 192.168.0.${ipnum}/24
nmcli connection modify ethStatic ipv6.method manual ipv6.addresses 2001:dead:1::${ipnum}

nmcli connection modify ethStatic ipv4.gateway 192.168.0.1
nmcli connection modify ethStatic ipv6.gateway 2001:dead:1::beef

nmcli connection modify ethStatic ipv4.dns 192.168.0.1
nmcli connection modify ethStatic ipv6.dns 2001:dead:1::beef

echo "ip address changed."

nmcli connection up ethStatic ifname eth0

