# Use this script to covert a .Msi file to Base64 Format

[Convert]::ToBase64String([IO.File]::ReadAllBytes("C:\path\to\your\installer.msi")) | clip