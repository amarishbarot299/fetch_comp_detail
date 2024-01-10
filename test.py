# import wmi

# computer = wmi.WMI()
# computer_info = computer.Win32_ComputerSystem()[0]
# os_info = computer.Win32_OperatingSystem()[0]
# proc_info = computer.Win32_Processor()[0]
# gpu_info = computer.Win32_VideoController()[0]

# os_name = os_info.Name.encode('utf-8').split(b'|')[0]
# os_version = ' '.join([os_info.Version, os_info.BuildNumber])
# system_ram = float(os_info.TotalVisibleMemorySize) / 1048576  # KB to GB

# print('OS Name: {0}'.format(os_name))
# print('OS Version: {0}'.format(os_version))
# print('CPU: {0}'.format(proc_info.Name))
# print('RAM: {0} GB'.format(system_ram))
# print('Graphics Card: {0}'.format(gpu_info.Name))



import psutil
import speedtest
import socket
import wmi

# Installed software list
def get_installed_software():
    software_list = []
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            pinfo = proc.as_dict(attrs=['pid', 'name'])
            software_list.append(pinfo['name'])
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return software_list

# Internet Speed
def get_internet_speed():
    st = speedtest.Speedtest()
    download_speed = st.download() / 1_000_000  # in Mbps
    upload_speed = st.upload() / 1_000_000  # in Mbps
    return f"Download Speed: {download_speed:.2f} Mbps, Upload Speed: {upload_speed:.2f} Mbps"

# Screen Resolution
def get_screen_resolution():
    try:
        from win32api import GetSystemMetrics
        width = GetSystemMetrics(0)
        height = GetSystemMetrics(1)
        return f"{width}x{height}"
    except ImportError:
        return "Cannot retrieve screen resolution on this system."

# # CPU Info
# def get_cpu_info():
#     cpu_info = {}
#     cpu_info['model'] = f"{psutil.cpu_brand()} {psutil.cpu_freq().current / 1000:.2f} GHz"
#     cpu_info['cores'] = psutil.cpu_count(logical=False)
#     cpu_info['threads'] = psutil.cpu_count(logical=True)
#     return cpu_info

# CPU Info
def get_cpu_info():
    cpu_info = {}
    cpu_info['model'] = None
    cpu_info['cores'] = psutil.cpu_count(logical=False)
    cpu_info['threads'] = psutil.cpu_count(logical=True)
    
    try:
        cpu_info_list = psutil.cpu_info()
        if cpu_info_list and isinstance(cpu_info_list, list):
            cpu_info['model'] = cpu_info_list[0]['brand_raw']
    except Exception as e:
        pass

    return cpu_info

# GPU Info (if available)
def get_gpu_info():
    try:
        w = wmi.WMI()
        gpu_info = w.Win32_VideoController()[0]
        return gpu_info.Name
    except Exception as e:
        return "No GPU information available."

# RAM Size
def get_ram_size():
    return f"{psutil.virtual_memory().total / (1024 ** 3):.2f} GB"

# Screen Size (assuming primary monitor)
def get_screen_size():
    try:
        from win32api import GetSystemMetrics
        diagonal_pixels = ((GetSystemMetrics(0) ** 2) + (GetSystemMetrics(1) ** 2)) ** 0.5
        diagonal_inches = diagonal_pixels / GetSystemMetrics(6)
        return f"{diagonal_inches:.1f} inches"
    except ImportError:
        return "Cannot retrieve screen size on this system."

# Mac Address
def get_mac_address():
    try:
        return ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) for elements in range(0, 2 * 6, 2)][::-1])
    except Exception as e:
        return "Cannot retrieve MAC address."

# Public IP Address
def get_public_ip():
    try:
        return socket.gethostbyname(socket.gethostname())
    except Exception as e:
        return "Cannot retrieve public IP address."

# Windows Version
def get_windows_version():
    try:
        w = wmi.WMI()
        os_info = w.Win32_OperatingSystem()[0]
        return f"{os_info.Caption} {os_info.BuildNumber}"
    except Exception as e:
        return "Cannot retrieve Windows version."

if __name__ == "__main__":
    print("Installed Software List:")
    for software in get_installed_software():
        print(f"- {software}")
    
    print("\nInternet Speed:", get_internet_speed())
    print("\nScreen Resolution:", get_screen_resolution())
    
    cpu_info = get_cpu_info()
    print(f"\nCPU Model: {cpu_info['model']}")
    print(f"No of Cores: {cpu_info['cores']}")
    print(f"No of Threads: {cpu_info['threads']}")
    
    print("\nGPU Model:", get_gpu_info())
    
    print("\nRAM Size:", get_ram_size())
    print("\nScreen Size:", get_screen_size())
    print("\nMAC Address:", get_mac_address())
    print("\nPublic IP Address:", get_public_ip())
    print("\nWindows Version:", get_windows_version())
