import cv2

# Load the video
video_path = "G:\BIkesh M 2081\ABC.mp4" # You can change the video path
cap = cv2.VideoCapture(video_path)

frame_count = 0
while True:
    ret, frame = cap.read()
    if not ret:
        break
    
    # Save a frame as an image
    frame_count += 1
    image_name = f"frame_{frame_count}.jpg"  # You can change the image format or name
    cv2.imwrite(image_name, frame)

cap.release()
print(f"Frames saved successfully.")
