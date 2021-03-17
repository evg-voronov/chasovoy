import time
import cv2
from mtcnn.mtcnn import MTCNN

class SessionWork:
    def __init__(self):
        self.start_work = 0
        self.start_break = 0
        self.time_break = 0
        self.time_work = 0

    def sit(self):
        if I.start_work == 0:
            self.start_work = time.time()
        self.start_break = 0
        self.time_break = 0

    def stand_up(self):
        if self.start_break == 0 and I.start_work != 0:
            self.start_break = time.time()

    def record(self):
        print(f"запись: {self.time_work}")
        SessionWork.__init__(self)

    def status(self):
        if self.start_work != 0 and self.start_break == 0:
            self.time_work = time.time() - self.start_work
        if self.start_break != 0:
            self.time_break = time.time() - self.start_break

    def show_faces(self, img, faces):
        if faces:
            bounding_box = faces[0]['box']
            cv2.rectangle(img,
                          (bounding_box[0], bounding_box[1]),
                          (bounding_box[0] + bounding_box[2], bounding_box[1] + bounding_box[3]),
                          (0, 0, 255),
                          2)
        cv2.imshow("I", img)



cap = cv2.VideoCapture(0)
detector = MTCNN()
min_time_work = 1800  # seconds. The minimum time after which the recording in Excel
min_time_break = 300  # seconds. The minimum break time after which the recording in Excel
I = SessionWork()

while cap.isOpened():
    _, img = cap.read()
    img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    faces = detector.detect_faces(img)
    if faces:
        I.sit()
    else:
        I.stand_up()
    I.status()
    print(f"start_work {I.start_work}")
    print(f"start_break {I.start_break}")
    print(f"time_work {I.time_work}")
    print(f"time_break {I.time_break}")
    print("-------------------")
    if I.time_work > min_time_work and I.time_break > min_time_break:
        I.record()
    I.show_faces(img, faces)
    cv2.waitKey(800)

cap.release()
cv2.destroyAllWindows()