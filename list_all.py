from win32com.shell import shell, shellcon
import os
import argparse

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

import json

def list_folders_and_files(base_ishellfolder):
    """Lists both folders and files in the given IShellFolder."""
    items = []
    for pidl in base_ishellfolder:
        name = base_ishellfolder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
        items.append(name)
    return items

def generate_structure(base_ishellfolder, depth=0, max_depth=None):
    """Recursively generate the folder and file structure."""
    if max_depth is not None and depth > max_depth:
        return None
    structure = {"folders": {}, "files": []}
    for pidl in base_ishellfolder:
        name = base_ishellfolder.GetDisplayNameOf(pidl, shellcon.SHGDN_NORMAL)
        attrs = base_ishellfolder.GetAttributesOf([pidl], shellcon.SFGAO_FOLDER)
        if attrs & shellcon.SFGAO_FOLDER:
            try:
                folder = base_ishellfolder.BindToObject(pidl, None, shell.IID_IShellFolder)
                structure["folders"][name] = generate_structure(folder, depth=depth+1, max_depth=max_depth)
            except:
                # If binding failed, it might be a special folder or inaccessible.
                # You can handle this case as you see fit.
                pass
        else:
            structure["files"].append(name)
    return structure

def save_structure_to_json(ishellfolder, output_file, max_depth=None):
    # save_structure_to_json(shell.SHGetDesktopFolder(), 'output.json')
    structure = generate_structure(ishellfolder, max_depth=max_depth)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(structure, f, ensure_ascii=False, indent=4)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--folder", help="folder path, e.g. \"This PC\\Apple iPhone\"")
    args = parser.parse_args()
    confirmed = False
    try:
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
                    selected_path ,_= navigate_and_select(shell.SHGetDesktopFolder(), [], is_first_time_returning=False)
            else:
                selected_path = ""  # Default to empty string if no items on the desktop

        print(f"Final selected path: {selected_path}")
        save_structure_to_json(get_folder_from_path(shell.SHGetDesktopFolder(), selected_path), 'output.json')

    except Exception as e:
        print("An error occurred:", e)
