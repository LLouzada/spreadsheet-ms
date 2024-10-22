# ANCP - SPREADSHEET MICROSERVICE

- Microservice to handle spreadsheet files manipulation.

## Installation

- Project Recommended Path (if changed, must be changed in the service too):
    - ~/projects/spreadsheet-ms
    - ~/miniconda3/ (default from conda installation)

- Install Conda (Miniconda):
    - `wget https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh`
    - `bash Miniconda3-latest-Linux-x86_64.sh`
    - `eval "$(/home/$USER/miniconda3/bin/conda shell.bash hook)"`
    - `conda init`

- create conda microservice env
    - `conda create -n spreadsheet-ms-env python=3.12 flask openpyxl`

- Make sure dependencies are installed
    - `conda install flask openpyxl`

- Activate / deactivate / remove **(if needed for testing)**
    - `conda activate spreadsheet-ms-env`
    - `conda deactivate`
    - `conda remove --name spreadsheet-ms-env --all`

- Execute the app **(just for testing, it is run by the service)**
    - cd ~/projects/spreadsheet-ms/app
    - python3 main.py

## Service Configuration

- Modify the service file to point to the correct path and environment
    - `~/projects/spreadsheet-ms/app/assets/files/ancp-spreadsheet-ms`
    - `ISPRODUCTION` -> define if it is production or not
    - make sure `MS_USER` is correct
    - make sure the PATHs are correct (if used recommended paths, there is no need to change)

- Copy the service file to /etc/init.d/ and make it executable
    - `sudo cp ~/projects/spreadsheet-ms/app/assets/files/ancp-spreadsheet-ms /etc/init.d/`
    - `sudo chmod +x /etc/init.d/ancp-spreadsheet-ms`

	### CentOS
	- `sudo chkconfig --add ancp-spreadsheet-ms`
	- `sudo service ancp-spreadsheet-ms start`
	- `sudo chkconfig ancp-spreadsheet-ms on` (boot enable)
	
	### Ubuntu
	- `sudo update-rc.d ancp-spreadsheet-ms defaults`
	- `sudo service ancp-spreadsheet-ms start`
  

## Service Log

  - `/var/log/ancp-spreadsheet-ms.log`
  - usage example: `tail -f /var/log/ancp-spreadsheet-ms.log`

