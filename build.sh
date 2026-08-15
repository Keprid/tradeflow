#!/usr/bin/env bash
set -e

apt-get update -y
apt-get install -y fonts-liberation || true
pip install -r requirements.txt
