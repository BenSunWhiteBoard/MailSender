# Send email

## Dependencies

Make sure:

1. On Windows platform
2. Python is installed
3. Microsoft Outlook Classic is installed and your email account has been setup correctly.

Python needs to install some extra packages, including:

```shell
pip install apscheduler extract-msg
```

## How to use

In powershell or gitbash:

```shell
# for help information
python ./send_email.py --help
```

```shell
# send test.msg file 5 times every 1 second, start from 2024-12-31 14:50:30 on local machine time
python .\send_email.py -num_run 5 -time_interval 1 -start_time "2024-12-31 14:50:30" -msg_file .\resources\test.msg

```