FROM python:3.10

RUN apt-get update -y
RUN apt-get install -y python-pip
RUN apt-get install -y cron

COPY . /app

WORKDIR /app

RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt

RUN chmod 444 main.py
RUN chmod 444 requirements.txt

# Copy hello-cron file to the cron.d directory
COPY work_hours_cron /etc/cron.d/work_hours_cron
 
# Give execution rights on the cron job
RUN chmod 0644 /etc/cron.d/work_hours_cron

# Apply cron job
RUN crontab /etc/cron.d/work_hours_cron
 
# Create the log file to be able to run tail
RUN touch /var/log/cron.log
 
# Run the command on container startup
CMD cron && tail -f /var/log/cron.log