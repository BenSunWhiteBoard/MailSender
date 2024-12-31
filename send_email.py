import os
import extract_msg
import win32com.client as win32


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


# Example usage
if __name__ == "__main__":
    msg_file = r"./resources/中文测试.msg"
    email_content = parse_msg_file(msg_file)
    if email_content:
        send_parsed_email(email_content)
