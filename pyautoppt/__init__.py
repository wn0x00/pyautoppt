import win32com.client


class Application(object):
    def __init__(self) -> None:
        self.application = win32com.client.Dispatch("PowerPoint.Application")
        self.application.Visible = 1


class Slide:
    def __init__(self) -> None:
        self.app = None
        self.presentation = None
        self.slides = None
        self.slide = None


class Slides:
    def __init__(self) -> None:
        self.app = None
        self.presentation = None
        self.slides = None

    def add_slide(self, index=0) -> Slide:
        slide_ref = Slide()
        # self.app = self.app
        # self.presentation = self.presentation
        # self.slides = self.slides

        # pptLayout = self.slides.Slides[1].CustomLayout
        # slide_ref.slide = self.slides.AddSlide(index, pptLayout)
        return slide_ref

    def add(self, index, layout):
        self.slides.add(index, layout)

    def find_by_slide_id(self):
        pass

    def get_enumerator(self):
        pass

    def insert_from_file(self):
        pass

    def paste(self):
        pass

    def range():
        pass


class Presentation:
    def __init__(self) -> None:
        self.app = None
        self.presentation = None

    @property
    def slides(self) -> Slides:
        slides_ref = Slides()
        slides_ref.app = self.app
        slides_ref.presentation = self.presentation
        slides_ref.slides = self.presentation.Slides
        return slides_ref

    def close(self):
        self.presentation.Close()
        self.app.Quit()

    def save(self) -> None:
        self.presentation.Save()

    def save_as(
        self,
        path,
        type,
    ):
        self.presentation.SaveAs(FileName=path, FileFormat=type, EmbedTrueTypeFonts=-2)


class Presentations(Application):
    """Presentations 对象, 可打开创建 pptx 文件"""

    def add(self) -> Presentation:
        presentation_ref = Presentation()
        presentation_ref.app = self.application
        presentation_ref.presentation = self.application.Presentations.Add()
        return presentation_ref

    def open(self, file_name: str) -> Presentation:
        self.application.Presentations.Open()
        presentation_ref = Presentation()
        presentation_ref.app = self.application
        presentation_ref.presentation = self.application.Presentations.Open(file_name)
        return presentation_ref


ppt = Presentations()

if __name__ == "__main__":
    # ppt = Presentations()
    # pre_ref = ppt.add()
    # pre_ref.slides.add(1,12)
    pass
