from frame_from_stream import Stream

Stream0 = Stream(0)
while True:
    frame = Stream0.read_frame_from_stream(look_frame=True)

