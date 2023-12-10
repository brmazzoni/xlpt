FROM python:alpine

WORKDIR xlpt

COPY . .

RUN pip install -r requirements.txt

ENTRYPOINT python xlpt/xlpt.py
