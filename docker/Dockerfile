# LibreOffice + python3-uno, so VBA compatibility probes can be run locally
# without installing LibreOffice on the host machine.
FROM debian:bookworm-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    # LibreOffice needs a writable HOME for its user profile.
    HOME=/tmp

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        python3 \
        python3-pip \
        python3-uno \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /work

COPY requirements.txt ./
RUN pip3 install --no-cache-dir --break-system-packages -r requirements.txt

COPY . .

CMD ["bash", "docker/run-class-test.sh"]
