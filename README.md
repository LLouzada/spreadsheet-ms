# ANCP - SPREADSHEET MICROSERVICE

- This microservice allows for the manipulation of Excel spreadsheet files, including the generation of custom documents with provided data.

## Installation

- **Recommended Project Path** (if changed, must be updated in the service file as well):
    - `~/projects/spreadsheet-ms`
    - `~/miniconda3/` (default from conda installation)

- **Install Conda (Miniconda)**:
    ```bash
    wget https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh
    bash Miniconda3-latest-Linux-x86_64.sh
    eval "$(/home/$USER/miniconda3/bin/conda shell.bash hook)"
    conda init
    ```

- **Create the microservice environment**:
    ```bash
    conda create -n spreadsheet-ms-env python=3.12 flask openpyxl
    ```

- **Environment Management**:
    - **Activate / Deactivate / Remove** (if needed for testing):
        ```bash
        conda activate spreadsheet-ms-env
        conda deactivate
        conda remove --name spreadsheet-ms-env --all
        ```

- **Run the app** (for testing only, it will be run by the service):
    ```bash
    cd ~/projects/spreadsheet-ms/app
    python3 main.py
    ```

## Service Configuration

- **Modify the service file** to point to the correct path and environment:
    - Define a ApiKey for the service in the `API_KEY` variable. This must be the same as the one send in the request headers.
    - `~/projects/spreadsheet-ms/app/assets/files/ancp-spreadsheet-ms`
    - `ISPRODUCTION` -> defines if it is in production or not.
    - Ensure `MS_USER` is correct and that the paths are correct (if using recommended paths, there is no need to change).

- **Copy the service file to `/etc/init.d/` and make it executable**:
    ```bash
    sudo cp ~/projects/spreadsheet-ms/app/assets/files/ancp-spreadsheet-ms /etc/init.d/
    sudo chmod +x /etc/init.d/ancp-spreadsheet-ms
    ```

	### CentOS
	```bash
	sudo chkconfig --add ancp-spreadsheet-ms
	sudo service ancp-spreadsheet-ms start
	sudo chkconfig ancp-spreadsheet-ms on  # enable at boot
	```
	
	### Ubuntu
	```bash
	sudo update-rc.d ancp-spreadsheet-ms defaults
	sudo service ancp-spreadsheet-ms start
	```
  
## Service Logs

- Log file location: `/var/log/ancp-spreadsheet-ms.log`
- Example usage: `tail -f /var/log/ancp-spreadsheet-ms.log`
