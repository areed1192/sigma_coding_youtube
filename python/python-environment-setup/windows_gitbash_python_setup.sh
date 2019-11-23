# USING PYTHON AND CONDA FROM GITBASH

Administrator rights.

# Adding Environment Variables to Windows to Run Python and Anaconda from GitBash

Python = 'C:\Users\Alex\Anaconda3'
Anaconda = 'C:\Users\Alex\Anaconda3\Scripts'

$ export PATH=~\Anaconda3:$PATH
$ export PATH=~\Anaconda3\Scripts:$PATH

$ py --version
$ python --version
$ conda --version

# Create a Virtual Environment from GitBash:

$ pip install virtualenv
$ py -m pip install virtualenv

$ python -m virtualenv my_envrionment_name
$ py -m virtualenv my_envrionment_name

Naviage to the 'my_environment_name' folder.

$ cd my_environment_name

# Activate Virtual Envrionment

$ . Scripts/activate - IF YOU'RE IN THE FOLDER
$ . my_environment_name/Scripts/activate - IF YOU'RE NOT IN THE FOLDER