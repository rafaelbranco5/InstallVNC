# InstallVNC
AutoIT Script to install TightVNC on a domain


This script installs TightVNC on a machine. By "spreading it trough AD it can install the software on almost all computers of a network.



Script only installs on Windows 7, 8 and 10.
Script ignores all other windows versions.

-> Validates whether a x64 or x32 bit version is needed and installs acordingly.

-> Validates whether it is already installed or not.

-> Creates a txt file on temp folder of machine with date and time of installation.

-> Creates a txt file with machine name and IP of machines where it wasn't installed on a shared folder.
