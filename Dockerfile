FROM python:3.12-alpine

RUN apk add --no-cache tzdata && \
    cp /usr/share/zoneinfo/Europe/Bratislava /etc/localtime && \
    echo "Europe/Bratislava" > /etc/timezone

ENV TZ=Europe/Bratislava

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY worklog_web.py .

RUN mkdir -p /data

EXPOSE 5000

CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "worklog_web:APP"]
