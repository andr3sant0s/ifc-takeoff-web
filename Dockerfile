
FROM python:3.10
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["gunicorn", "-b", "0.0.0.0:10000", "--timeout", "600", "--workers", "1", "--threads", "2", "app:app"]
