FROM python:3.10.12
RUN apt-get update; apt-get install python3-pip -y ;  pip install pandas openpyxl