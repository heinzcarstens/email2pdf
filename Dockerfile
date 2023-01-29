FROM ubuntu

MAINTAINER Nicolas Goldack <nicolas-goldack@live.de>

WORKDIR /email2pdf/

# Build-time metadata as defined at http://label-schema.org
ARG BUILD_DATE
ARG VCS_REF
ARG VERSION
LABEL io.ntec.email2pdf.build-date=$BUILD_DATE \
      io.ntec.email2pdf.name="Email2Pdf" \
      io.ntec.email2pdf.description="Email to pdf downloader written in python" \
      io.ntec.email2pdf.url="https://github.com/ntec-io/Email2Pdf" \
      io.ntec.email2pdf.vcs-ref=$VCS_REF \
      io.ntec.email2pdf.vcs-url="https://github.com/ntec-io/Email2Pdf" \
      io.ntec.email2pdf.vendor="ntec.io" \
      io.ntec.email2pdf.version=$VERSION \
      io.ntec.email2pdf.schema-version="1.0"

RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 \
    wkhtmltopdf \
    wget \
    python3-pip

RUN wget "https://chilkatdownload.com/9.5.0.93/chilkat-9.5.0-python-3.8-x86_64-linux.tar.gz" --no-check-certificate
    
RUN cd chilkat-9.5.0-python-3.8-x86_64-linux && python3 installChilkat.py

ADD . .

RUN pip3 install -r requirements.txt


CMD ["python3", "-u", "email2pdf.py"]
