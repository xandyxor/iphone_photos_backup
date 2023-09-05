from win32com.shell import shell, shellcon
import os
import argparse
import shutil
import pythoncom

def list_folders(base_ishellfolder):
    return [base_ishellfolder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL) for pidl in base_ishellfolder]

def select_folder_from_list(folder_list,is_first_time_returning=False):
    while True:
        if not is_first_time_returning:
            selection = input("Enter the number, 'y', or 'n': ").strip()
        else:
            selection = input("Enter 'y', or 'n': ").strip()

        if selection.isdigit() and not is_first_time_returning:
            selected_number = int(selection)
            if 1 <= selected_number <= len(folder_list):
                return folder_list[selected_number - 1]
            else:
                print("Invalid number. Please select again.")
        elif selection in ['y', 'n']:
            return selection
        else:
            print("Invalid selection. Please try again.")

def navigate_and_select(base_ishellfolder, path_parts=[], is_first_time_returning=True):
    folders = list_folders(base_ishellfolder)
    folder_path = "\\".join(path_parts) if path_parts else "Desktop"

    print(f'Exploring folder: {folder_path}')
    print("================")
    print()
    for i, folder_name in enumerate(folders, 1):
        print(f"({i}). {folder_name}")
    print()
    print("================")
    print("Options:")
    print("Enter folder number to navigate")
    print("Confirm path (y)")
    # if not is_first_time_returning:  # Only show "Go back to previous folder" option if not the first time returning
    print("Go back to previous folder (n)")
    print("================")
    print(f'Exploring folder: {folder_path}')


    if is_first_time_returning:
        selection = select_folder_from_list(folders,is_first_time_returning=True)
    else:
        selection = select_folder_from_list(folders,is_first_time_returning=False)

    if selection == 'y':
        return "\\".join(path_parts), True
    elif selection == 'n':
        if is_first_time_returning:
            path_parts = []  # 清空路徑，導向桌面層級
            return navigate_and_select(shell.SHGetDesktopFolder(), is_first_time_returning=False)
        else:
            if path_parts:
                path_parts.pop()
                parent_ishellfolder = get_folder_from_path(shell.SHGetDesktopFolder(), "\\".join(path_parts)) if path_parts else None
                return navigate_and_select(parent_ishellfolder, path_parts, is_first_time_returning=False)
    else:
        path_parts.append(selection)
        selected_folder = get_folder_from_path(base_ishellfolder, selection)
        return navigate_and_select(selected_folder, path_parts, is_first_time_returning=False)


def get_folder_from_path(base_ishellfolder, path):
    first_part, *rest = path.split("\\")
    for pidl in base_ishellfolder:
        if base_ishellfolder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL) == first_part:
            folder = base_ishellfolder.BindToObject(pidl, None, shell.IID_IShellFolder)
            return folder if not rest else get_folder_from_path(folder, "\\".join(rest))
    return None


def get_pidl_from_virtual_path(virtual_path):
    parts = virtual_path.split('\\')
    current_folder = shell.SHGetDesktopFolder()

    for part in parts:
        found = False
        for pidl in current_folder:
            if current_folder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL) == part:
                current_folder = current_folder.BindToObject(pidl, None, shell.IID_IShellFolder)
                found = True
                break
        if not found:
            return None
    return shell.SHGetIDListFromObject(current_folder)

def copy_item_to_destination(source_path, destination_path):
    pidl_source = get_pidl_from_virtual_path(source_path)
    pidl_dest = shell.SHILCreateFromPath(destination_path, 0)[0]

    if not pidl_source:
        raise Exception("無法獲取源路徑的PIDL")

    si_source = shell.SHCreateItemFromIDList(pidl_source, shell.IID_IShellItem)
    si_dest = shell.SHCreateItemFromIDList(pidl_dest, shell.IID_IShellItem)

    # 檢查來源與目標是否相同
    source_display_name = si_source.GetDisplayName(shellcon.SIGDN_FILESYSPATH)
    dest_display_name = si_dest.GetDisplayName(shellcon.SIGDN_FILESYSPATH)
    if source_display_name == dest_display_name:
        raise Exception("目的資料夾與來源資料夾相同")

    # 更進一步的檢查，確保目標資料夾不是源資料夾的子資料夾
    if dest_display_name.startswith(source_display_name):
        raise Exception("目的資料夾是來源資料夾的子資料夾")

    pfo = pythoncom.CoCreateInstance(shell.CLSID_FileOperation, None, pythoncom.CLSCTX_ALL, shell.IID_IFileOperation)
    pfo.SetOperationFlags(shellcon.FOF_NOCONFIRMATION)
    pfo.CopyItem(si_source, si_dest, None)
    pfo.PerformOperations()

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--folder", help="folder path, e.g. \"This PC\\Apple iPhone\"")
    args = parser.parse_args()
    confirmed = False
  
    if args.folder:
        selected_path = args.folder
    else:
        desktop_folder = shell.SHGetDesktopFolder()
        desktop_folder_items = list_folders(desktop_folder)
        if desktop_folder_items:
            first_item_on_desktop = desktop_folder_items[0]
            default_path = f"{first_item_on_desktop}\\Apple iPhone\\Internal Storage\\DCIM2"
            try:
                selected_folder = get_folder_from_path(shell.SHGetDesktopFolder(), default_path)
                # navigate_and_select(selected_folder, default_path.split("\\"))
                selected_path, confirmed = navigate_and_select(selected_folder, default_path.split("\\"))
                # selected_path = default_path
            except Exception as e:
                print(f"An error occurred while navigating to the default path: {e}")
                print("Switching to manual mode...")
                selected_path ,confirmed= navigate_and_select(shell.SHGetDesktopFolder(), [], is_first_time_returning=False)
        else:
            selected_path = ""  # Default to empty string if no items on the desktop

    print(f"confirmed: {confirmed},Final selected path: {selected_path}")
    copy_item_to_destination(selected_path, "D:\\ANDY\\test\\")