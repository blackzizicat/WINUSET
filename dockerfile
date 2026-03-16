FROM python:3.12-slim

WORKDIR /app

RUN pip install --no-cache-dir openpyxl requests

COPY main.py auth.py ./

CMD ["python", "main.py"]
