import os
import sys
import subprocess
import site

def install_geopandas():
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "geopandas"])
        return True
    except Exception as e:
        print("Failed to install GeoPandas:", e)
        return False

def get_geopandas_path():
    try:
        from pathlib import Path
        site_packages = [Path(p) for p in sys.path if p.endswith("site-packages")]
        for sp in site_packages:
            geopandas_path = sp / "geopandas"
            if geopandas_path.is_dir():
                return str(geopandas_path.parent)
        return None
    except Exception as e:
        print("Failed to get GeoPandas path:", e)
        return None

try:
    import geopandas as gpd
except ImportError:
    installed = install_geopandas()
    if installed:
        geopandas_path = get_geopandas_path()
        if geopandas_path is not None:
            sys.path.append(geopandas_path)
            try:
                import geopandas as gpd
            except ImportError:
                print("Failed to import GeoPandas even after installation.")
            else:
                print("GeoPandas successfully imported after installation.")
        else:
            print("Failed to find GeoPandas path.")
    else:
        print("GeoPandas installation failed.")
else:
    print("GeoPandas successfully imported.")
