
import urllib.request, json, sys

# Read token from git credential or use PAT stored somewhere
# Try to get it from git config
import subprocess
result = subprocess.run([r'C:\Program Files\Git\cmd\git.exe', 'config', '--global', 'github.token'], capture_output=True, text=True)
token = result.stdout.strip()
if not token:
    # Try Windows credential manager via git credential-manager
    result2 = subprocess.run([r'C:\Program Files\Git\mingw64in\git-credential-manager.exe', 'get'], 
                              input=b"protocol=https
host=github.com
", capture_output=True)
    print("No token found, trying other method...")
    sys.exit(1)
print(f"Token starts with: {token[:8]}...")
