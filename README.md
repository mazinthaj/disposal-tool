## Install:
- pip install pandas
- pip install tk
- pip install pyzbar
- pip install opencv-python
- pip install openpyxl
- pip install pillow
- pip3 install pillow-heif

## Description

When a customer does not collect their packages, the packages are destroyed after a certain period of time. The following process is carried out:

- Before the package is destroyed, a photo of the item is taken.
- After the package is destroyed, a second photo is taken, showing either the destroyed package or the destruction process.

Previously, this process was done manually. After the entire process, the user would:

- Go through each image.
- Find the waybill number.
- Rename the image files to `XXXXXXXXX_before` and `XXXXXXXXX_after`.
- Type the image location into an Excel file.

### Before:
![image](https://github.com/user-attachments/assets/aef694e5-3de9-4ea4-a59c-02bbf5810c16)
### After:
![image](https://github.com/user-attachments/assets/b2d7536a-654e-46a4-b456-ab640cb5665d)

## allData.py

- The waybill number in the package must be displayed in both the "before" and "after" photos.



## allData2.py

- No need for the waybill number in the "after" photo.
- Once a barcode is detected in the "before" photo, the system automatically marks the next image as the "after" photo.

### Photo of the application
![image](https://github.com/user-attachments/assets/28153081-18fc-425f-b807-214a85bf37cb)


## Logic Behind the Image

- Example filename: `20241114_070001641_iOS.heic`
- The numbers in the filename represent the date and time the image was taken.



## Conditions

- Both "before" and "after" photos must exist.
- No photo should be taken unless the "after" photo follows the corresponding "before" photo.
- The image must be taken on an iPhone, which generates images with the `.HEIC` file extension.
