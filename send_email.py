import argparse
import pythoncom
import os
import extract_msg
import win32com.client as win32
from apscheduler.schedulers.background import BackgroundScheduler
from datetime import datetime, timedelta
from threading import Event

# Event to block the program until all jobs are done
job_done = Event()


def parse_msg_file(file_path):
    """
    Parse a .msg email file and extract its contents.

    Args:
        file_path (str): Path to the .msg file.

    Returns:
        dict: Parsed email contents including subject, body, recipients, and attachments.
    """
    try:
        msg = extract_msg.Message(file_path)
        msg_data = {
            "subject": msg.subject,
            "body": msg.body,
            "to": msg.to,
            "attachments": []
        }

        # Save attachments to a temporary directory
        attachment_dir = "attachments"
        os.makedirs(attachment_dir, exist_ok=True)
        for attachment in msg.attachments:
            attachment_path = os.path.join(attachment_dir, attachment.longFilename)
            with open(attachment_path, "wb") as f:
                f.write(attachment.data)
            msg_data["attachments"].append(attachment_path)

        return msg_data
    except Exception as e:
        print(f"Failed to parse .msg file: {e}")
        return None


def send_parsed_email(email_data):
    """
    Send an email using Outlook with parsed data.

    Args:
        email_data (dict): Parsed email contents.
    """
    try:
        # Initialize COM library
        pythoncom.CoInitialize()

        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = email_data.get("subject", "No Subject")
        mail.Body = email_data.get("body", "No Body")
        # Extract only the email address
        recipient = email_data.get("to", "").strip()
        if '<' in recipient and '>' in recipient:
            recipient = recipient[recipient.find('<') + 1:recipient.find('>')].strip()
        mail.To = recipient

        print("Email Content(Mail Object):")
        print(f"To: {mail.To}")
        print(f"Subject: {mail.Subject}")
        print(f"Body: {mail.Body}")
        if email_data.get("attachments"):
            print("Attachments:")
            for attachment in email_data["attachments"]:
                print(f"- {attachment}")

        # Add attachments if any
        for attachment in email_data.get("attachments", []):
            absolute_path = os.path.abspath(attachment)
            if not os.path.exists(absolute_path):
                print(f"Attachment path does not exist: {absolute_path}")
                continue  # Skip this attachment
            mail.Attachments.Add(absolute_path)

        mail.Send()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")


def send_parsed_email_wrapper(run_count: int, email_data):
    print(f"Run {run_count} function called at {datetime.now()}!")
    send_parsed_email(email_data)
    # Signal that the job is done
    if run_count == num_run:
        # Signal the last run
        job_done.set()
    else:
        print(f"{num_run - run_count} jobs left")


def main(num_run, time_interval, start_time, email_content):
    scheduler = BackgroundScheduler()

    # Parse start_time into a datetime object
    start_datetime = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S")
    now_datetime = datetime.now()

    # Calculate the delay in seconds
    delay_seconds = (start_datetime - now_datetime).total_seconds()
    if not delay_seconds > 0:
        print(f"Target time{start_datetime} is in the past. Now is {now_datetime}. Cannot schedule.")
        return

        # Schedule the function num_run times with time_interval seconds interval
    for i in range(1, num_run + 1):
        run_time = start_datetime + timedelta(seconds=time_interval * (i - 1))
        scheduler.add_job(send_parsed_email_wrapper, 'date', run_date=run_time, args=[i, email_content])
        print(f"Scheduled run {i} at {run_time}")

    # Start the scheduler
    scheduler.start()

    # Wait for all jobs to complete
    job_done.wait()

    # Shutdown the scheduler
    scheduler.shutdown()
    print("All jobs completed. Scheduler shut down.")


# Example usage
if __name__ == "__main__":
    # Parse command-line arguments
    parser = argparse.ArgumentParser(
        description="Schedule a function num_run times with time_interval seconds interval starting at start_time.")
    parser.add_argument("-num_run", type=int, required=True, help="Number of times to run the function")
    parser.add_argument("-time_interval", type=float, required=True, help="Interval in seconds between runs")
    parser.add_argument("-start_time", type=str, required=True,
                        help="Start time for the first run in 'YYYY-MM-DD HH:MM:SS' format")
    parser.add_argument("-msg_file", type=str, required=True, default=None,
                        help="Path to the message file for scheduled sending")
    args = parser.parse_args()

    # Extract arguments
    num_run = args.num_run
    time_interval = args.time_interval
    start_time = args.start_time
    msg_file = args.msg_file

    email_content = parse_msg_file(msg_file)

    # Run the scheduler
    main(num_run, time_interval, start_time, email_content)
