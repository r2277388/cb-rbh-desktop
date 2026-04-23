#!/bin/bash

# Create the .ssh directory if it doesn't exist
mkdir -p ~/.ssh

# Copy SSH keys from Google Drive to the .ssh directory
cp /content/drive/MyDrive/ssh/id_ed25519 ~/.ssh/
cp /content/drive/MyDrive/ssh/id_ed25519.pub ~/.ssh/

# Set the correct permissions for the private key
chmod 600 ~/.ssh/id_ed25519

# Start the SSH agent and add the key
eval "$(ssh-agent -s)"
ssh-add ~/.ssh/id_ed25519

# Optional: Add GitHub to known hosts
ssh-keyscan -t ed25519 github.com >> ~/.ssh/known_hosts

echo "SSH setup complete. You can now use git commands with GitHub."
