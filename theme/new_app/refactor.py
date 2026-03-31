import os

os.chdir(r"c:\Users\micks\Documents\GitHub\Data-Analyse\theme\new_app")

def read(f):
    with open(f, "r", encoding="utf-8") as file: return file.read()
    
def write(f, content):
    with open(f, "w", encoding="utf-8") as file: file.write(content)

# 1. engine.py
code_engine = read("models.py") + "\n" + read("database.py") + "\n" + read("quality_engine.py") + "\n" + read("centrale_engine.py")
write("engine.py", code_engine)

# 2. home.py
code_home = read("auth.py") + "\n" + read("home.py")
code_home = code_home.replace("from models import User", "from engine import User")
code_home = code_home.replace("import database", "import engine as database")
write("home_new.py", code_home)

# 3. dashboard.py
code_dash = read("dashboard.py")
code_dash = code_dash.replace("import database", "import engine as database")
write("dashboard_new.py", code_dash)

# 4. ana.py
code_ana = read("loading.py") + "\n" + read("main.py")
code_ana = code_ana.replace("from loading import LoadingFrame", "")
code_ana = code_ana.replace("from auth import AuthFrame", "from home import AuthFrame")
code_ana = code_ana.replace("import database", "import engine as database")
code_ana = code_ana.replace("from quality_engine import QualityEngine", "from engine import QualityEngine")
code_ana = code_ana.replace("from centrale_engine import CentraleEngine", "from engine import CentraleEngine")
write("ana.py", code_ana)

# Cleanup
for f in ["models.py", "database.py", "quality_engine.py", "centrale_engine.py", "auth.py", "loading.py", "main.py", "home.py", "dashboard.py"]:
    if os.path.exists(f): os.remove(f)

# Rename
os.rename("home_new.py", "home.py")
os.rename("dashboard_new.py", "dashboard.py")

print("Herstructurering succesvol voltooid! ✅")
