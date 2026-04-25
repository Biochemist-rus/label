import socket

ip = "172.16.102.80"
port = 9100

line1 = "Привет".encode("utf-8")
line2 = "УФ лампа".encode("utf-8")

zpl = (
    b"^XA\r\n"
    b"^CI28\r\n"
    b"^FO50,40\r\n"
    b"^A0N,34,34\r\n"
    b"^FD" + line1 + b"^FS\r\n"
    b"^FO50,85\r\n"
    b"^A0N,34,34\r\n"
    b"^FD" + line2 + b"^FS\r\n"
    b"^XZ\r\n"
)

with socket.create_connection((ip, port), timeout=10) as s:
    s.sendall(zpl)