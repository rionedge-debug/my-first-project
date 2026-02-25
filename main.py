from datetime import datetime


def print_datetime():
    now = datetime.now()
    print(now.strftime("%Y-%m-%d %H:%M:%S"))


if __name__ == "__main__":
    print_datetime()
