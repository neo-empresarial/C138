FROM python:3.9-slim

COPY . /app

WORKDIR /app

RUN pip install -r requirements.txt

EXPOSE 8501

RUN mkdir ~/.streamlit

CMD ["streamlit", "run", "script.py"]
