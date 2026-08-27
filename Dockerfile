FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY sync.py .
# Run every 15 minutes via a simple shell loop
CMD ["sh", "-c", "while true; do python sync.py; sleep 900; done"]
