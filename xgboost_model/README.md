# How to start this process:

1) if on c-drive, use venv
2) if on colab - vscode tunnel
    0) currently using colab default libraries
    1) open a new colab and run:
        !pip install -U git+https://github.com/amitness/colab-connect.git
    2) then run:
        from colabconnect import colabconnect
        colabconnect()
    3) you'll receive in the output a link with a code ... follow link and input.
    4) Open VSCode on your laptop and open the command prompt
    5) Select: 'Remote-Tunnels: Connect to Tunnel' to connect to colab
    6) github
    7) choose open tunnel
3) Setup space
    1) to use git, run 
        bash colab_ssh_set.sh
    2) here are a couple modules to upgrade
        bash colab_setup.sh

###########################################
### Colab Session Start
Because colab forgets your data each time you log off... a script was created to access the github key 
& a numpy upgrade and a xgboost upgrade are necessary. Please update using the following script when using colab:

python colab_init.py

###########################################
### FYI: 
Tensorflow only supports python 3.11 please use the venv to begin using the appropriate python version.

## Getting Started

### Github Info
Using the following email: r2277388+github@gmail.com
Github Username: r2277388
SSH: git@github.com:r2277388/cb-rbh-desktop.git
HTTPS: https://github.com/r2277388/cb-rbh-desktop.git

### Pulling and Pushing Changes

To begin using this app, first activate the virtual environment:

```bash
source venv/Scripts/activate
git push origin main
git pull origin main
```

### Setting up the Virtual Environment

To begin using this app, first activate the virtual environment:

```bash
source venv/Scripts/activate
```

### Workflow to pip install another module

Make sure you are in the venv:

```bash
source venv/Scripts/activate
pip install <>
pip freeze > requirements.txt
```

### Instructions to Clone this Repository

1. **Clone the Repository**:
    ```bash
    git clone <repository-url>
    cd <repository-name>
    ```

2. **Create a Virtual Environment**:
    ```bash
    python -m venv venv
    ```

3. **Activate the Virtual Environment**:
    ```bash
    source venv/Scripts/activate
    ```

4. **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

### Running the Program

To run the program, use:

```bash
python main.py
```