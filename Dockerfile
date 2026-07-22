FROM python:3.11-slim

WORKDIR /auditor

RUN pip install --no-cache-dir pyyaml

COPY . /auditor
RUN mkdir -p reports

ENTRYPOINT ["./audit.sh"]
CMD ["--help"]
