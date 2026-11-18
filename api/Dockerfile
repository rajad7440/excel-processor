FROM python:3.11-slim
WORKDIR /app
COPY requirements-api.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt
COPY processor.py api.py ./
ENV PORT=8000
EXPOSE 8000
CMD ["sh", "-c", "uvicorn api:app --host 0.0.0.0 --port ${PORT}"]
