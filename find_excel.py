import os

def find_file(root_dir, target_filename):
    for root, dirs, files in os.walk(root_dir):
        if target_filename in files:
            return os.path.join(root, target_filename)
    return None

root = r"C:\Users\hangs\OneDrive"
# Let's search for a part of the filename that might be stable
target = "채권선도 거래.xlsx"
path = find_file(root, target)

if path:
    print(f"FOUND: {path}")
else:
    print("NOT FOUND")
