
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))

from config_manager import ConfigManager

def patch():
    cm = ConfigManager()
    try:
        scripts = cm.load_scripts_config()
        changed = False
        
        # Patch Kvadra -> RIR Energo
        if 'kvadra' in scripts:
            if scripts['kvadra']['name'] == "Kvadra":
                print("Patching kvadra name...")
                scripts['kvadra']['name'] = "RIR Energo"
                changed = True
                
        # Patch RVK Kvadra -> RVK RIR Energo (Optional, but logical if the company renamed. 
        # But per instruction I will strict to Kvadra->RIR, but "RVK Kvadra" allows users to differentiate.
        # I will leave RVK Kvadra alone unless asked, to avoid confusion, or maybe just check if user wants it.)
        
        if changed:
            cm._save_json(cm.scripts_path, scripts)
            print("Successfully updated script name in active configuration.")
        else:
            print("Configuration already has the new name or script not found.")
            
    except Exception as e:
        print(f"Error patching: {e}")

if __name__ == "__main__":
    patch()
