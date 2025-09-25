import subprocess
import logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
def run_script(script_name):
   try:
       logging.info(f"Running {script_name}...")
       subprocess.run(["python", script_name], check=True)
       logging.info(f"{script_name} completed successfully. ✅")
   except subprocess.CalledProcessError as e:
       logging.error(f"{script_name} failed with exit code {e.returncode}")
   except Exception as e:
       logging.error(f"Unexpected error running {script_name}: {e}")
if __name__ == "__main__":
   scripts = [
       "Tag.py",
       "Back.py",
       "patch.py"
   ]
   for script in scripts:
       run_script(script)
   logging.info("All scripts executed. Reports are stored locally in the current folder.")