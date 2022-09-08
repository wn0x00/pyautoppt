import unittest
import time

from pyautoppt import ppt


class Base(unittest.TestCase):
    def test_add_presentation(self):
        ppt_ref = ppt.add()
        time.sleep(3)
        ppt_ref.close()

    def test_add_slide(self):
        pre_ref = ppt.add()
        time.sleep(3)
        pre_ref.slides.add(1,12)


if __name__ == "__main__":
    unittest.main()
