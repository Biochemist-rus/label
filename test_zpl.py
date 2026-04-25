import socket

ip = "172.16.102.80"  # свой IP
port = 9100

zpl = b"""
^XA
^FO50,50
^ADN,36,20
^FDHELLO ZPL^FS
^XZ
"""

with socket.create_connection((ip, port), timeout=5) as s:
    s.sendall(zpl)