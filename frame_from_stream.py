import cv2
from datetime import datetime


class Stream:
    def __init__(self, id_camera):
        self.current_second = datetime.now().strftime("%S")
        self.cap = cv2.VideoCapture(id_camera)

    def read_frame_from_stream(self, look_frame: bool):
        """read one frame in second"""
        if self.cap.isOpened():
            _, frame = self.cap.read()
            if datetime.now().strftime("%S") == self.current_second:
                return self.read_frame_from_stream(look_frame)
            else:
                self.current_second = datetime.now().strftime("%S")
                if look_frame:
                    cv2.imshow('Look', frame)
                    cv2.waitKey(1)
                return frame
        else:
            self.cap.release()
            cv2.destroyAllWindows()
            print("camera not available")
            exit()








