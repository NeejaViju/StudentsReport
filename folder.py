import os
from PIL import Image, ImageDraw, ImageFont

student_data = {
    "PPP001": "Mohamed Hasir",
    "PPP002": "Ganesh Kumar R",
    "PPP003": "Deepa N",
    "PPP004": "Nt. Nallathayammal",
    "PPP005": "Prasanth Govindaraj",
    "PPP006": "Murali T",
    "PPP007": "LEEMAN THOMAS",
    "PPP008": "Vimal Nadarajan",
    "PPP009": "Saravanan Selvam",
    "PPP010": "Srinivasan SR",
    "PPP011": "David Raj",
    "PPP012": "Yogesh Kumar JG",
    "PPP013": "Aravindhan Selvaraj",
    "PPP014": "Naveen Bromiyo A R",
    "PPP015": "Kalai Selvi",
    "PPP016": "Madhan Karthick",
    "PPP017": "Pavithra Selvaraj",
    "PPP018": "Sindhu Laheri Uthaya Surian",
    "PPP019": "Nalina Athinamilagi",
    "PPP020": "Nithya Naveen",
    "PPF001": "Ranjitha",
    "PPF002": "Suganthi Ramaraj",
    "PPF004": "Swathipriya",
    "PPF005": "Jumana",
    "PPF006": "Indira Priyadharshini",
    "PPF007": "Riyas ahamed J"
}

weeks = 52  # Number of weeks

for student_id, student_name in student_data.items():
    student_folder = os.path.join("Students", f"{student_id} - {student_name}")
    
    for week in range(1, weeks + 1):
        week_folder = os.path.join(student_folder, f"Week{week:02}")
        if not os.path.exists(week_folder):
            os.makedirs(week_folder)
            print(f"Created folder for {student_name} - Week{week:02}")
        else:
            print(f"Folder already exists for {student_name} - Week{week:02}, skipping...")

        # Create and save the sample.png image in the week folder
        image = Image.new('RGB', (200, 200), color='white')
        draw = ImageDraw.Draw(image)
        font = ImageFont.load_default()  # Use default font

        draw.text((10, 10), f"Week {week:02}", fill='black', font=font)
        draw.text((10, 30), f"Student: {student_name}", fill='black', font=font)

        image.save(os.path.join(week_folder, 'sample.png'))
        print(f"Created sample.png for {student_name} - Week{week:04}")

print("Folder creation and image generation completed!")
