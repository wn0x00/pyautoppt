import win32com.client


class Application(object):
    def __init__(self) -> None:
        self.application = win32com.client.Dispatch("PowerPoint.Application")
        self.application.Visible = 1


if __name__ == "__main__":
    app = Application()
    print(app.application.Version)
