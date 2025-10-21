FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Run migrations and collect static files
RUN python manage.py collectstatic --noinput

EXPOSE 8000

# Use shell form to run multiple commands
CMD python manage.py migrate && python manage.py runserver 0.0.0.0:8000
