import cv2
from face_encoding import Face  # Import the Face class

# Initialize the face recognizer
custom_encoder = Face()

# Load face encodings from the dataset folder (ensure this path exists and has images)
custom_encoder.load_encoding_images("Dataset/")

# Set for unique names to avoid duplicates in attendance
unique_names = set()

# Open attendance log file in append mode
attendance_log = open("attendance_log.txt", "a")

# Start capturing video from webcam
cap = cv2.VideoCapture(0)

while True:
    ret, frame = cap.read()
    
    if not ret:
        print("Failed to capture frame from camera. Exiting...")
        break

    # Detect faces and get the recognized names
    face_areas, recognized_names = custom_encoder.detect_known_faces(frame)

    # Display the recognized faces and update attendance
    for face_coords, found_name in zip(face_areas, recognized_names):
        t, r, b, l = face_coords  # Extract face coordinates
        
        # Add the recognized name to attendance log if not already added
        if found_name not in unique_names:
            unique_names.add(found_name)
            attendance_log.write(found_name + "\n")

        # Draw a rectangle around the face and display the name
        cv2.putText(frame, found_name, (l, t - 10), cv2.FONT_HERSHEY_DUPLEX, 1, (0, 0, 200), 2)
        cv2.rectangle(frame, (l, t), (r, b), (0, 0, 200), 4)

    # Display the result frame
    cv2.imshow("Face Recognition Attendance System", frame)

    # Break the loop if 'q' is pressed
    if cv2.waitKey(3) & 0xFF == ord('q'):
        break

# Release the webcam and close windows
cap.release()
cv2.destroyAllWindows()

# Close attendance log file
attendance_log.close()
