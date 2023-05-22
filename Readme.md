# Fibocom L8x0 Connect for Windows

![](./screenshot/screen01.png)

## Run

All script **_must_** be run as administrator

- `connect.cmd`: Connect and monitoring
- `monitor.cmd`: Connection monitoring without connect

## Setup

#### APN

Edit `scripts/main.ps1` to configure your carrier APN, APN_USER and APN_PASS

#### Preferred bands

Find `AT+XACT=` in `scripts/main.ps1` and edit command to your needs