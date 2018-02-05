import logging
try:
    import winreg
except ImportError:
    import _winreg as winreg

from .refresh_icons import refresh_icons
from ._version import get_versions

__version__ = get_versions()['version']
del get_versions


__all__ = ['refresh_icons', 'delete_tree', 'get_class', 'register_fileclass',
           'register_extension', 'delete_class', 'unregister_fileclass',
           'unregister_extension']


def delete_tree(root_key, sub_key):
    '''
    Recursively delete subkey (including all subkeys and values).

    Parameters
    ----------
    root_key : already open key, or one of the predefined HKEY_* constants
    sub_key : str
        Subkey to delete.
    '''
    with winreg.OpenKey(root_key, sub_key) as open_key:
        key_info = dict(zip(('key_count', 'value_count', 'modified'),
                            winreg.QueryInfoKey(open_key)))

        for i in xrange(0, key_info['key_count']):
            sub_key_i = winreg.EnumKey(open_key, 0)
            try:
                winreg.DeleteKey(open_key, sub_key_i)
                logging.debug(r'Removed `%s`', sub_key_i)
            except Exception:
                # Subkey has child subkeys.
                delete_tree(open_key, sub_key_i)

        winreg.DeleteKey(open_key, '')
        logging.debug('Removed %s', sub_key)


def get_class(name, all_users=False):
    '''
    Parameters
    ----------
    name : str
        Name of class registry key.
    all_users : bool, optional
        If ``True``, register fileclass for all users.

        Otherwise, only register fileclass for current user.
    '''
    root_name = 'HKEY_CURRENT_USER' if not all_users else 'HKEY_CLASSES_ROOT'
    root_key = getattr(winreg, root_name)

    # Create fileclass registry key.
    if not all_users:
        fileclass_path = 'Software\Classes\%s' % name
    else:
        fileclass_path = name
    return winreg.OpenKey(root_key, fileclass_path)


def register_fileclass(name, details, all_users=False, description=None,
                       overwrite=False):
    '''
    Create file class.

    Fileclass **SHOULD** define:
     - Description
     - Default verb (e.g., "open")
     - Verb description(s) (e.g., `"Open with MicroDrop"`)
     - Verb command(s) (e.g., `<full path>\microdrop-3 open "%*"`)

    See `here <https://msdn.microsoft.com/en-us/library/windows/desktop/cc144158%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396>`
    for more info.

    Parameters
    ----------
    name : str
        Name of fileclass
    details : dict
    all_users : bool, optional
        If ``True``, register fileclass for all users.

        Otherwise, only register fileclass for current user.
    description : str, optional
        Fileclass description.
    overwrite : bool, optional
        Overwrite existing fileclass.


    .. versionchanged:: 0.1.1
        Fix: set value of fileclass key to specified description.
    '''
    root_name = 'HKEY_CURRENT_USER' if not all_users else 'HKEY_CLASSES_ROOT'
    root_key = getattr(winreg, root_name)

    # Create fileclass registry key.
    if not all_users:
        fileclass_path = 'Software\Classes\%s' % name
    else:
        fileclass_path = name
    try:
        fileclass_key = winreg.OpenKey(root_key, fileclass_path)
        if overwrite:
            delete_tree(root_key, fileclass_path)
        else:
            raise KeyError('Fileclass already exists for `%s\%s`', root_name,
                           fileclass_path)
    except WindowsError:
        pass
    logging.debug('Create fileclass key for `%s\%s`', root_name,
                  fileclass_path)
    fileclass_key = winreg.CreateKey(root_key, fileclass_path)
    if description is not None:
        winreg.SetValue(root_key, fileclass_path, winreg.REG_SZ, description)

    for name_i, value_i in details.iteritems():
        if value_i is not None:
            winreg.SetValue(fileclass_key, name_i, winreg.REG_SZ, value_i)


def register_extension(extension, fileclass, details=None, all_users=False,
                       overwrite=False):
    '''
    Parameters
    ----------
    extension : str
        File extension (e.g., ``.exe``).
    fileclass : str
        Fileclass to associate with file extension.
    details : dict
    all_users : bool, optional
        If ``True``, register fileclass for all users.

        Otherwise, only register fileclass for current user.
    overwrite : bool, optional
        Overwrite existing file extension.
    '''
    root_name = 'HKEY_CURRENT_USER' if not all_users else 'HKEY_CLASSES_ROOT'
    root_key = getattr(winreg, root_name)

    # Create fileclass registry key.
    if not all_users:
        extension_path = 'Software\Classes\%s' % extension
    else:
        extension_path = extension
    try:
        extension_key = winreg.OpenKey(root_key, extension_path)
        if overwrite:
            delete_tree(root_key, extension_path)
        else:
            raise KeyError('Extension already exists for `%s\%s`' %
                           (root_name, extension_path))
    except WindowsError:
        pass
    logging.debug('Create extension key for `%s\%s`', root_name,
                  extension_path)
    extension_key = winreg.CreateKey(root_key, extension_path)
    winreg.SetValue(root_key, extension_path, winreg.REG_SZ, fileclass)

    if details is not None:
        for name_i, value_i in details.iteritems():
            if value_i is not None:
                winreg.SetValueEx(extension_key, name_i, None, winreg.REG_SZ,
                                  value_i)


def delete_class(name, all_users=False):
    '''
    Parameters
    ----------
    name : str
        Name of class key to delete.
    all_users : bool, optional
        If ``True``, register fileclass for all users.

        Otherwise, only register fileclass for current user.
    '''
    root_name = 'HKEY_CURRENT_USER' if not all_users else 'HKEY_CLASSES_ROOT'
    root_key = getattr(winreg, root_name)

    # Create fileclass registry key.
    if not all_users:
        fileclass_path = 'Software\Classes\%s' % name
    else:
        fileclass_path = name

    delete_tree(root_key, fileclass_path)


def unregister_fileclass(name, all_users=False):
    '''
    Parameters
    ----------
    name : str
        Name of fileclass to unregister.
    all_users : bool, optional
        If ``True``, register fileclass for all users.

        Otherwise, only register fileclass for current user.
    '''
    delete_class(name, all_users=all_users)


def unregister_extension(name, all_users=False, with_fileclass=True):
    if with_fileclass:
        try:
            with get_class(name, all_users=all_users) as key:
                _, fileclass, _ = winreg.EnumValue(key, 0)
            unregister_fileclass(fileclass, all_users=all_users)
        except WindowsError:
            pass
    delete_class(name, all_users=all_users)
