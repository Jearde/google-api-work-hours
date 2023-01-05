FROM python:3.10

RUN apt-get update -y
RUN apt-get install -y cron

# COPY . /app
COPY requirements.txt /app/requirements.txt
COPY ./work_hours /app/work_hours
COPY ./run_python.sh /app/run_python.sh

WORKDIR /app

RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt

RUN chmod 444 work_hours/main.py
RUN chmod 444 requirements.txt
RUN chmod +x /app/run_python.sh

# Copy hello-cron file to the cron.d directory
COPY work_hours_cron /etc/cron.d/work_hours_cron
 
# Give execution rights on the cron job
RUN chmod 0644 /etc/cron.d/work_hours_cron

# Apply cron job
RUN crontab /etc/cron.d/work_hours_cron
 
# Create the log file to be able to run tail
RUN touch /var/log/cron.log

EXPOSE 8088

# Run the command on container startup
CMD /app/run_python.sh && cron && tail -f /var/log/cron.log