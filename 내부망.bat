netsh interface ipv4 set address name="Wi-Fi" source=static address=00.000.0.000 mask=255.255.255.0 gateway=00.000.0.0
netsh dnsclient add dnsservers name="Wi-Fi" index=1 address=000.000.000.0
netsh dnsclient add dnsservers name="Wi-Fi" index=2 address=
