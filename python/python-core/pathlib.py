import pathlib

# Let's create a basic path to this file.
print(pathlib.Path(__file__))

"""
    Path classes are divided between pure paths,
    which provide purely computational operations
    without I/O, and concrete paths, which inherit
    from pure paths but also provide I/O operations.
"""

# Let's take at a Windows Path and a Pure Windows Path.
windows_path = pathlib.WindowsPath(__file__)
windows_path_pure = pathlib.PureWindowsPath(__file__)
print("PurePath: " + str(windows_path_pure))
print("Path: " + str(windows_path))

"""
    Here's a case where this might be useful. You 
    are on Linux but youneed to manipulate a windows
    path.

    Let's explore the operations we can do with a path.
"""

# We can grab the parts.
print(windows_path.parts)

# There is also a property called `parent`. This will allow
# us to get the directory above.
print(windows_path.parent)

# and you can attach one call on top of the other.
print(windows_path.parent.parent)

# But when you're doing that you really should use `parents`.
# This way you can just select how you want to go up. For
# example, if you want to recreate the code up above just
# do this:
print(windows_path.parents[1])

# I want to create a folder in this director, so let's define the path.
data_folder = pathlib.Path(__file__).parents[0].joinpath('data')

# Let's check to see if it already exists first, by using the `exists()`
# method.
if not data_folder.exists():
    data_folder.mkdir()

# Let's also add a file inside of it, but this time I want to raise
# an error if it already exists. We don't to overwrite the file!
data_folder.joinpath('my_file.txt').touch(exist_ok=True)

# Lets add some more objects
data_folder.joinpath('my_file.csv').touch(exist_ok=True)
data_folder.joinpath('my_file.json').touch(exist_ok=True)
data_folder.joinpath('my_folder').mkdir(exist_ok=True)

# From here lets iterate through the folder.
for file_obj in data_folder.iterdir():

    # Let's ask some questions about each object.
    print("+"*80)
    print(file_obj)
    print("Are you a directory? " + str(file_obj.is_dir()))
    print("Are you a file? " + str(file_obj.is_file()))
    print("Are you a symbolic link? " + str(file_obj.is_symlink()))
    print("Are you a relative of the `Data` Folder? " + str(file_obj.is_relative_to(data_folder)))
    print("File Drive is: " + str(file_obj.drive))
    print("File Stem is: " + str(file_obj.stem))
    print("File Anchor is: " + str(file_obj.anchor))
    print("File Name is: " + str(file_obj.name))
    print("File Suffix is: " + str(file_obj.suffix))

# Here are some useful methods.

# Grab the Current Working Directory.
print("Current Working Directory: " + str(pathlib.Path.cwd()))

# Grab the File Path of THIS script.
print("Full File Path of the current Script: " + str(pathlib.Path(__file__)))

# Grab the System home path.
print("Home Path is: " + str(pathlib.Path.home()))

# `absolute` takes a partial path and makes it a full path.
print("My Partial Path looks like this before `absolute`: " + str(pathlib.Path("data")))
print("My Partial Path looks like this after `absolute`: " + str(pathlib.Path("data").absolute()))

# `resolve` will do things like remove `..` or change windows path to unix paths and vice versa.
print(pathlib.Path('docs/../setup.py').resolve())
print(pathlib.Path("C:/Users/Alex/OneDrive/Growth - Tutorial Videos/Lessons - Python/Lessons - Pathlib/using_pathlib.py").resolve())

# Final thing is `stat()` which can be used to get certain statistics about the file.
print(pathlib.Path(__file__).stat().st_size)
print(pathlib.Path(__file__).stat().st_mtime)
