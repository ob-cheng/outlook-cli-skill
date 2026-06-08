"""Outlook folder management."""

DEFAULT_FOLDERS = {
    'inbox': 6,
    'sent': 5,
    'sent items': 5,
    'drafts': 16,
    'deleted': 3,
    'deleted items': 3,
    'outbox': 4,
    'junk': 23,
    'junk email': 23,
    'calendar': 9,
    'contacts': 10,
    'tasks': 13,
    'notes': 12,
}


def list_all_folders(namespace) -> list[dict]:
    """Recursively list all folders in the mailbox.

    Returns:
        List of dicts with keys: name, path, level, count, is_store
    """
    folders = []

    def traverse_folders(folder_collection, path="", level=0):
        for i in range(1, folder_collection.Count + 1):
            try:
                folder = folder_collection.Item(i)
                folder_path = f"{path}/{folder.Name}" if path else folder.Name
                folders.append({
                    'name': folder.Name,
                    'path': folder_path,
                    'level': level,
                    'count': folder.Items.Count if hasattr(folder, 'Items') else 0
                })
                if folder.Folders.Count > 0:
                    traverse_folders(folder.Folders, folder_path, level + 1)
            except Exception:
                continue

    for i in range(1, namespace.Folders.Count + 1):
        try:
            store = namespace.Folders.Item(i)
            folders.append({
                'name': store.Name,
                'path': store.Name,
                'level': 0,
                'count': 0,
                'is_store': True
            })
            traverse_folders(store.Folders, store.Name, 1)
        except Exception:
            continue

    return folders


def find_folder_by_name(namespace, folder_name):
    """Find a folder by name (case-insensitive).

    Args:
        namespace: Outlook MAPI namespace
        folder_name: Folder name to find

    Returns:
        Folder object or None
    """
    folder_lower = folder_name.lower().strip()

    if folder_lower in DEFAULT_FOLDERS:
        folder_id = DEFAULT_FOLDERS[folder_lower]
        if folder_id > 0:
            try:
                return namespace.GetDefaultFolder(folder_id)
            except Exception:
                pass

    def search_folders(folder_collection, target_name):
        for i in range(1, folder_collection.Count + 1):
            try:
                folder = folder_collection.Item(i)
                if folder.Name.lower() == target_name:
                    return folder
                if folder.Folders.Count > 0:
                    result = search_folders(folder.Folders, target_name)
                    if result:
                        return result
            except Exception:
                continue
        return None

    for i in range(1, namespace.Folders.Count + 1):
        try:
            store = namespace.Folders.Item(i)
            if store.Name.lower() == folder_lower:
                return store
            result = search_folders(store.Folders, folder_lower)
            if result:
                return result
        except Exception:
            continue

    return None


def find_folder_by_path(namespace, folder_path):
    """Find a folder by full path (e.g., 'Account/Inbox/Subfolder').

    Args:
        namespace: Outlook MAPI namespace
        folder_path: Full path to folder

    Returns:
        Folder object or None
    """
    parts = [p.strip() for p in folder_path.split('/') if p.strip()]
    if not parts:
        return None

    current = None
    for i in range(1, namespace.Folders.Count + 1):
        try:
            store = namespace.Folders.Item(i)
            if store.Name.lower() == parts[0].lower():
                current = store
                break
        except Exception:
            continue

    if not current:
        return find_folder_by_name(namespace, parts[-1])

    for part in parts[1:]:
        found = False
        for j in range(1, current.Folders.Count + 1):
            try:
                subfolder = current.Folders.Item(j)
                if subfolder.Name.lower() == part.lower():
                    current = subfolder
                    found = True
                    break
            except Exception:
                continue
        if not found:
            return None

    return current
