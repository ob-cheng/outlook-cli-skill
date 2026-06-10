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


def list_all_folders(namespace, use_cache: bool = True) -> list[dict]:
    """Recursively list all folders in the mailbox.

    Args:
        namespace: Outlook MAPI namespace.
        use_cache: If True, return cached tree when fingerprint matches.
                   Set False to force a full namespace walk.

    Returns:
        List of dicts with keys: name, path, level, is_store
    """
    # Check cache first (fast path — fingerprint only, no tree walk)
    if use_cache:
        from .folder_cache import load_cache, save_cache
        cached = load_cache(namespace)
        if cached is not None:
            return cached

    folders = _walk_folders(namespace)

    # Save to cache for next call
    if use_cache:
        try:
            from .folder_cache import save_cache
            save_cache(folders, namespace)
        except Exception:
            pass  # cache write failure shouldn't break the command

    return folders


def _walk_folders(namespace) -> list[dict]:
    """Recursively walk the namespace folder tree (no caching)."""
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


def find_folder_across_stores(namespace, folder_name: str):
    """Find a folder by name across ALL stores, skipping the GetDefaultFolder shortcut.

    Unlike find_folder_by_name, this always searches every store recursively instead of
    falling back to GetDefaultFolder() for well-known folder names. Needed for multi-account
    setups where calendar/tasks/notes data lives in a non-default store.

    Args:
        namespace: Outlook MAPI namespace
        folder_name: Folder name to find (case-insensitive)

    Returns:
        Folder object from the first matching store, or None
    """
    target = folder_name.lower().strip()

    def recursive_search(folder_collection):
        for i in range(1, folder_collection.Count + 1):
            try:
                folder = folder_collection.Item(i)
                if folder.Name.lower() == target:
                    return folder
                if folder.Folders.Count > 0:
                    result = recursive_search(folder.Folders)
                    if result:
                        return result
            except Exception:
                continue
        return None

    for i in range(1, namespace.Folders.Count + 1):
        try:
            store = namespace.Folders.Item(i)
            if store.Name.lower() == target:
                return store
            result = recursive_search(store.Folders)
            if result:
                return result
        except Exception:
            continue

    # Fallback to GetDefaultFolder if nothing found across stores
    default_map = {
        'calendar': 9, 'tasks': 13, 'notes': 12,
        'inbox': 6, 'sent items': 5, 'drafts': 16,
        'deleted items': 3, 'junk email': 23, 'contacts': 10,
    }
    if target in default_map:
        try:
            return namespace.GetDefaultFolder(default_map[target])
        except Exception:
            pass
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
