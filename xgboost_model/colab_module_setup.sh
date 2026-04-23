#!/bin/bash

# Upgrade pip
pip install --upgrade pip

# Install compatible versions of numpy, pandas, and xgboost
pip install "numpy==1.26.4" "pandas==2.2.2" --upgrade xgboost

pip install ipykernel==5.5.6

# Optionally install other dependencies from requirements.txt
# pip install -r /content/drive/MyDrive/code_xgboost/requirements.txt