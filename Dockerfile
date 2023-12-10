FROM python:alpine

WORKDIR xlpt

COPY . .

RUN pip install -r requirements.txt

CMD python xlpt/xlpt.py
