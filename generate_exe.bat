python -m pip install virtualenv==20.0.15
python -m virtualenv --no-setuptools venv
call venv\Scripts\activate
pip3 install -r requirements.txt
pyinstaller -F filter.py
move dist\filter.exe filter.exe
