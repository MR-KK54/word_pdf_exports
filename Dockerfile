FROM python:3.12-slim

# Word files are paginated against installed fonts.  These font families cover
# the common Arial/Calibri-compatible metrics used by Office documents.
RUN apt-get update && apt-get install -y --no-install-recommends \
    fontconfig fonts-crosextra-carlito fonts-dejavu-core fonts-liberation \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f

WORKDIR /opt/render/project/src
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PYTHONUNBUFFERED=1
CMD ["gunicorn", "--workers", "1", "--threads", "4", "--timeout", "120", "run_web:application"]
