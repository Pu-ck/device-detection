import datetime
import time
import string
import win32file
import win32gui
import win32con
import msvcrt
import configparser


class Device:
    def __init__(self, type, drive, space):
        self.type = type
        self.drive = drive
        self.free_space = round((space[2] * space[1] * space[0]) * 9.31 * 1e-10, 2)
        self.total_space = round((space[3] * space[1] * space[0]) * 9.31 * 1e-10, 2)


# Detecting devices - creating new Device class object and adding it to a list if given drive is not already occupied
def devices_detection(drive_letter_list, devices_list, drives_list):
    for drive_letter in drive_letter_list:
        if drive_letter not in drives_list:
            if (
                win32file.GetDriveType(drive_letter + ":/") == win32file.DRIVE_FIXED
                and drive_fixed_detection == True
            ):
                device = Device(
                    "Fixed drive",
                    drive_letter,
                    win32file.GetDiskFreeSpace(drive_letter + ":/"),
                )
                devices_list.append(device)
            elif (
                win32file.GetDriveType(drive_letter + ":/") == win32file.DRIVE_REMOVABLE
                and drive_removable_detection == True
            ):
                device = Device(
                    "Removable drive",
                    drive_letter,
                    win32file.GetDiskFreeSpace(drive_letter + ":/"),
                )
                devices_list.append(device)
            elif (
                win32file.GetDriveType(drive_letter + ":/") == win32file.DRIVE_REMOTE
                and drive_remote_detection == True
            ):
                device = Device(
                    "Remote drive",
                    drive_letter,
                    win32file.GetDiskFreeSpace(drive_letter + ":/"),
                )
                devices_list.append(device)
            elif (
                win32file.GetDriveType(drive_letter + ":/") == win32file.DRIVE_CDROM
                and cdrom_detection == True
            ):
                device = Device(
                    "CD-ROM",
                    drive_letter,
                    win32file.GetDiskFreeSpace(drive_letter + ":/"),
                )
                devices_list.append(device)
            elif (
                win32file.GetDriveType(drive_letter + ":/") == win32file.DRIVE_RAMDISK
                and drive_ram_detection == True
            ):
                device = Device(
                    "RAM drive",
                    drive_letter,
                    win32file.GetDiskFreeSpace(drive_letter + ":/"),
                )
                devices_list.append(device)
            elif win32file.GetDriveType(drive_letter + ":/") == win32file.DRIVE_UNKNOWN:
                device = Device(
                    "Unknown drive",
                    drive_letter,
                    win32file.GetDiskFreeSpace(drive_letter + ":/"),
                )
                devices_list.append(device)


# Print information about connected devices and their connection time, save it to .txt file
def devices_connection(drives_list):
    for device in devices_list:
        if (
            device.drive not in drives_list
            and win32file.GetDriveType(device.drive + ":/") != 1
        ):
            print("Device connected at: ", end="")
            print(now.strftime("%Y-%m-%d %H:%M:%S"))
            print("Device type: " + device.type)
            print("Device drive: " + device.drive)
            print("Device free space: " + str(device.free_space) + "GB")
            print("Device total space: " + str(device.total_space) + "GB\n")
            drives_list.append(device.drive)

            with open("connections_list.txt", "a") as connections_list:
                connections_list.write("Device connected at: ")
                connections_list.write(now.strftime("%Y-%m-%d %H:%M:%S") + "\n")
                connections_list.write("Device type: " + device.type + "\n")
                connections_list.write("Device drive: " + device.drive + "\n")
                connections_list.write(
                    "Device free space: " + str(device.free_space) + "GB\n"
                )
                connections_list.write(
                    "Device total space: " + str(device.total_space) + "GB\n\n"
                )

            global show_list
            show_list = True


# Print information about disconnected devices and their disconnection time, save it to .txt file
def devices_removal(drives_list, devices_list):
    for drive in drives_list:
        if win32file.GetDriveType(drive + ":/") == 1:
            print("Device disconnected at: ", end="")
            print(now.strftime("%Y-%m-%d %H:%M:%S"))

            for device in devices_list:
                if drive == device.drive:
                    print("Device type: " + device.type)
                    break

            print("Device drive: " + drive)
            print("Device free space: " + str(device.free_space) + "GB")
            print("Device total space: " + str(device.total_space) + "GB\n")
            drives_list.remove(drive)
            devices_list.remove(device)

            with open("connections_list.txt", "a") as connections_list:
                connections_list.write("Device disconnected at: ")
                connections_list.write(now.strftime("%Y-%m-%d %H:%M:%S") + "\n")
                connections_list.write("Device type: " + device.type + "\n")
                connections_list.write("Device drive: " + drive + "\n")
                connections_list.write(
                    "Device free space: " + str(device.free_space) + "GB\n"
                )
                connections_list.write(
                    "Device total space: " + str(device.total_space) + "GB\n\n"
                )

            global show_list
            show_list = True


# Show list of currently connected devices
def show_devices(devices_list):
    print("Number of currently connected devices: " + str(len(devices_list)))
    for device in devices_list:
        print("- " + device.type, end="")
        print(" at drive: " + device.drive, end="")
        print(
            ". Free space: "
            + str(device.free_space)
            + "GB, total space: "
            + str(device.total_space)
            + "GB"
        )
    print()


# Update .ini file with device type detection settings
def update_settings():
    try:
        with open("settings.ini", "w") as config_updated:
            config.set(
                "DetectionSettings", "drive_fixed_detection", str(drive_fixed_detection)
            )
            config.set(
                "DetectionSettings",
                "drive_removable_detection",
                str(drive_removable_detection),
            )
            config.set(
                "DetectionSettings",
                "drive_remote_detection",
                str(drive_remote_detection),
            )
            config.set("DetectionSettings", "cdrom_detection", str(cdrom_detection))
            config.set(
                "DetectionSettings", "drive_ram_detection", str(drive_ram_detection)
            )
            config.set(
                "DetectionSettings", "drive_ram_detection", str(drive_ram_detection)
            )
            config.set("DetectionSettings", "hide_window", str(hide_window))
            config.write(config_updated)
    except FileNotFoundError:
        print("File does not exist")


if __name__ == "__main__":

    # List of 26 letters to match devices drives, list of Device class objects and list of currently occupied drives
    drive_letter_list = string.ascii_uppercase[:26]
    devices_list = []
    drives_list = []

    show_list = False

    # Read device type detection settings from .ini file
    try:
        config = configparser.ConfigParser()
        config.read("settings.ini")
        config.options("DetectionSettings")
        drive_fixed_detection = config.getboolean(
            "DetectionSettings", "drive_fixed_detection"
        )
        drive_removable_detection = config.getboolean(
            "DetectionSettings", "drive_removable_detection"
        )
        drive_remote_detection = config.getboolean(
            "DetectionSettings", "drive_remote_detection"
        )
        cdrom_detection = config.getboolean("DetectionSettings", "cdrom_detection")
        drive_ram_detection = config.getboolean(
            "DetectionSettings", "drive_ram_detection"
        )
        hide_window = config.getboolean("DetectionSettings", "hide_window")
    except FileNotFoundError:
        print("Settings file does not exist - settings set to default")
        drive_fixed_detection = True
        drive_removable_detection = True
        drive_remote_detection = True
        cdrom_detection = True
        drive_ram_detection = True
        hide_window = False

    # Creating .txt file to store informations about connections
    try:
        connections_list = open("connections_list.txt", "x")
    except FileExistsError:
        connections_list = open("connections_list.txt", "a")

    # Hiding console window if hide_window in settings.ini is set to True
    if hide_window == True:
        hide = win32gui.GetForegroundWindow()
        win32gui.ShowWindow(hide, win32con.SW_HIDE)

    print("[1] Show connection list")
    print("[2] Delete connection list")
    print("[3] Settings")
    print("[4] Exit program\n")

    while True:

        # Simple menu - not necessarily needed (same as update_settings() function), mainly for sake of testing stuff
        if msvcrt.kbhit():

            menu = msvcrt.getch()
            menu = str(menu)

            # Show connections list stored in .txt file
            if menu == "b'1'":
                try:
                    connections_list = open("connections_list.txt", "r")
                    print("Connections list: \n")
                    print(connections_list.read())
                    connections_list.close()
                except FileNotFoundError:
                    print("Can't show connections list - ", end="")
                    print("file does not exist or the list is empty\n")
            # Delete .txt file with connections list
            elif menu == "b'2'":
                connections_list.close()
                try:
                    print("Deleting connections list...")
                    win32file.DeleteFile("connections_list.txt")
                    print("Connections list deleted\n")
                except BaseException:
                    print("Can't delete file - ", end="")
                    print("file does not exist or the list is empty\n")
            # Print current device type detection settings stored in .ini file
            elif menu == "b'3'":
                if drive_fixed_detection == True:
                    print("[1] Fixed drives detection: enabled")
                elif drive_fixed_detection == False:
                    print("[1] Fixed drives detection: disabled")

                if drive_removable_detection == True:
                    print("[2] Removable drives detection: enabled")
                elif drive_removable_detection == False:
                    print("[2] Removable drives detection: disabled")

                if drive_remote_detection == True:
                    print("[3] Remote drives detection: enabled")
                elif drive_remote_detection == False:
                    print("[3] Remote drives detection: disabled")

                if cdrom_detection == True:
                    print("[4] CD-ROM detection: enabled")
                elif cdrom_detection == False:
                    print("[4] CD-ROM detection: disabled")

                if drive_ram_detection == True:
                    print("[5] RAM drives detection: enabled")
                elif drive_ram_detection == False:
                    print("[5] RAM drives detection: disabled")

                settings = msvcrt.getch()
                settings = str(settings)
                # Change device type detection settings and store them in .ini file
                if settings == "b'1'" and drive_fixed_detection == True:
                    print("Detecting fixed drives disabled\n")
                    drive_fixed_detection = False
                    update_settings()
                elif settings == "b'1'" and drive_fixed_detection == False:
                    print("Detecting fixed drives enabled\n")
                    drive_fixed_detection = True
                    update_settings()
                elif settings == "b'2'" and drive_removable_detection == True:
                    print("Detecting removable drives disabled\n")
                    drive_removable_detection = False
                    update_settings()
                elif settings == "b'2'" and drive_removable_detection == False:
                    print("Detecting removable drives enabled\n")
                    drive_removable_detection = True
                    update_settings()
                elif settings == "b'3'" and drive_remote_detection == True:
                    print("Detecting remote drives disabled\n")
                    drive_remote_detection = False
                    update_settings()
                elif settings == "b'3'" and drive_remote_detection == False:
                    print("Detecting remote drives enabled\n")
                    drive_remote_detection = True
                    update_settings()
                elif settings == "b'4'" and cdrom_detection == True:
                    print("Detecting CD-ROMs disabled\n")
                    cdrom_detection = False
                    update_settings()
                elif settings == "b'4'" and cdrom_detection == False:
                    print("Detecting CD-ROMs drives enabled\n")
                    cdrom_detection = True
                    update_settings()
                elif settings == "b'5'" and drive_ram_detection == True:
                    print("Detecting RAM drives disabled\n")
                    drive_ram_detection = False
                    update_settings()
                elif settings == "b'5'" and drive_ram_detection == False:
                    print("Detecting RAM drives enabled\n")
                    drive_ram_detection = True
                    update_settings()
                else:
                    pass
            # Exit
            elif menu == "b'4'":
                print("Exiting program\n")
                break
            else:
                pass

        now = datetime.datetime.now()
        devices_connection(drives_list)
        devices_removal(drives_list, devices_list)
        devices_detection(drive_letter_list, devices_list, drives_list)

        if show_list == True:
            show_devices(devices_list)
            show_list = False

        time.sleep(0.5)
