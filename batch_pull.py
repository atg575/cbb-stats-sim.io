import subprocess
import sys

def run_master_scraper():
    # List of all your scraping scripts verbatim
    scripts = [
        "kenpom_puller.py",
        "height_puller.py",
        "evan-miya-ratings.py",
        "espn_bpi_puller.py",
        "elo_ratings.py",
        "cbbref_stat_puller.py",
        "compute_stats.py"
    ]

    print("Starting Master CBB Scraper...\n")

    for script in scripts:
        print(f"{'='*40}")
        print(f"Running {script}...")
        print(f"{'='*40}")
        
        try:
            # sys.executable ensures the script uses the exact same Python 
            # environment you are currently running the master script with.
            subprocess.run([sys.executable, script], check=True)
            print(f"Successfully finished {script}.\n")
            
        except subprocess.CalledProcessError as e:
            print(f"ERROR: {script} failed with exit code {e.returncode}.\n")
        except FileNotFoundError:
            print(f"ERROR: Could not find '{script}'. Check the spelling and folder location.\n")

    print("All scraping scripts have finished executing!")
    print("Master stats recomputed in data/stats_master.json — no Excel recalculation needed!")

if __name__ == "__main__":
    run_master_scraper()