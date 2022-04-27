# stagehand tool by patrick

import sys
import wmi
import click
import ctypes
import re
import shutil
from time import sleep
import os
import subprocess
import json
import pyautogui
from pathlib import Path

# constants

CFG_PATH = Path(r"\\192.168.1.100\map_as_y\Brink\Customers\FFC\Staging")
IP_PREFIX = "192.168.128."
TZ_LIST = [
    "Eastern Standard Time",
    "Central Standard Time",
    "Mountain Standard Time",
    "Pacific Standard Time"
]
DESKTOP = Path(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))
#checks if site is a NRO and then it'll auto delete shortcuts and reboot terminal
NRO = False

VIG_TERM = ""

# kitchen terminal IDs by name
# helps sort kitchens without opening config files (b/c NAS lags)
KITCHEN_IDS = {
    "Grill": 21,
    "Grill 2": 22,
    "Make": 23,
    "Custard": 24,
    "DT Expo": 25,
    "DT Grill": 26,
    "DT Grill 2": 27,
    "DT Make": 28,
    "Expo": 29,
    "DT Expo 2": 30
}


# set progress flags
class Flags(object):
    """
    Flags class is used to keep track of staging progress and use settings across terminals.
    Flags are stored as files in the site's directory on NAS, so some settings only have to be set once, ex. timezone.

    >>> flags = Flags(site_no, term_name)
    >>> flags["flagname"] = value   # set flag specific to terminal
    >>> flags["_flagname"] = value  # set flag for all terms
    """

    def __init__(self, site, term_id):
        self.site = site
        self.site_path = CFG_PATH / site  # site path

        if not self.site_path.is_dir():
            raise ValueError(f"Can't find {site}")
        
        self.path = self.site_path / str(term_id) / ".flags"
        self.path.mkdir(parents=True, exist_ok=True)

        self.gpath = self.site_path / ".flags"
        self.gpath.mkdir(exist_ok=True)
    
    def __getitem__(self, name):
        path = self.path
        if name[0] == "_":
            path = self.gpath
            name = name[1:]
        flag_path = path / (name + ".txt")

        if not flag_path.is_file():
            return None
        
        return json.load(open(str(flag_path), "r"))
    
    def __setitem__(self, name, val):
        path = self.path
        if name[0] == "_":
            path = self.gpath
            name = name[1:]
        flag_path = path / (name + ".txt")

        with open(str(flag_path), "w") as file:
            file.write(json.dumps(val))


# menu

def menu(prompt, options, force=False):
    # shows menu in terminal
    
    click.echo(prompt)
    i = 0
    for opt in options:
        i += 1
        click.echo(f"  {i}. " + opt)

    selected = click.prompt(">>", default=None if force else "skip")

    try:
        selected = int(selected) - 1
    except ValueError:
        selected = -1

    if not (0 <= selected < len(options)):
        if force:
            return menu(prompt, options, True)
        else:
            return None

    return selected


# UTIL

def quickedit(enabled=True):  # This is a patch to the console quickedit system that sometimes hangs
    kernel32 = ctypes.windll.kernel32
    if enabled:
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x40|0x100))
        print("Console Quick Edit Enabled")
    else:
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-10), (0x4|0x80|0x20|0x2|0x10|0x1|0x00|0x100))
        print("Console Quick Edit Disabled")


def set_static_ip(ip, subnet_mask, gateway, dns=("8.8.8.8", "8.8.4.4")):
    nic_configs = wmi.WMI().Win32_NetworkAdapterConfiguration(IPEnabled=True)

    # First network adaptor
    nic = nic_configs[0]

    # Set IP address, subnetmask and default gateway
    # Note: EnableStatic() and SetGateways() methods require *lists* of values to be passed
    nic.EnableStatic(IPAddress=[ip], SubnetMask=[subnet_mask])
    nic.SetGateways(DefaultIPGateway=[gateway])
    nic.SetDNSServerSearchOrder(list(dns))


def rename_pc(name):
    c = wmi.WMI()

    for system in c.Win32_ComputerSystem():
        system.Rename(name)


def eblvd(pc_name, path=r"\\192.168.1.100\map_as_y\Brink\Installers\eBLVD-ra-146D80.exe"):
    p = subprocess.Popen(path)  # open eblvd setup exe
    sleep(2)

    pyautogui.FAILSAFE = False

    pyautogui.press("enter")
    sleep(2)
    pyautogui.typewrite(pc_name, interval=0.1)
    pyautogui.press(["tab", "tab", "tab", "tab", "space", "tab", "enter"], interval=0.2)
    sleep(5)
    pyautogui.press(['enter'])
    sleep(1)


# check store number

def get_site_info(site_no, sort=True):
    print("searching.", end="", flush=True)
    
    site_dir = CFG_PATH / site_no
    if not site_dir.is_dir():
        return None

    registers = []
    kitchens = []

    for f in site_dir.glob("*/*.cfg"):
        print(".", end="", flush=True)
        
        fn = f.stem.lower()
        if "_" in fn:
            continue
        if fn == "register":
            registers.append(f.parent.name)
        elif fn == "kitchen":
            kitchens.append(f.parent.name)

    def k_key(v):
        # Key function used to sort KVS's by term no.

        print(".", end="", flush=True)

        # option 1: if KVS name is common (in KITCHEN_IDS set at top of file), use common term no.
        if v in KITCHEN_IDS:
            return KITCHEN_IDS[v]
        
        # option 2: open config file and read terminal number directly.
        # NAS doesn't like when you do this a lot.
        cfg_path = str(CFG_PATH / site_no / v / "Kitchen.cfg")
        with open(cfg_path) as file:
            m = re.search(r'TerminalNumber="(\d+)"', file.read())
            
            if m is None:
                click.echo("Couldn't find terminal number", err=True)
                return 0

            return int(m.group(1))

    if sort:  # only spend time sorting kitchens if user is actually staging a kitchen
        kitchens = sorted(kitchens, key=k_key)
    

    return {
        "registers": sorted(registers),  # basic sort works for register names R1 - R9
        "kitchens": kitchens
    }

# get vig agents

def get_vig(site_no):
    # return list of vigilix installer paths in site directory

    site_dir = CFG_PATH / site_no
    if not site_dir.is_dir():
        return None

    agents = []
    for f in site_dir.glob("*.exe"):
        fn = f.stem
        if "_" in fn:
            continue

        agents.append((fn, f))  # return both the file names (for menu) and full path (for exec)

    return agents

def get_latest_brinkadminpanel():
    source=r"\\192.168.1.100\map_as_y\Brink\Customers\FFC\pdill\BrinkAdminPanel.exe"
    destination=r"C:\Brink\Pos\BrinkAdminPanel.exe"
    try:
        shutil.copy(source, destination)
        print("Latest Brink Admin Panel retrieved.")
    except shutil.SameFileError:
        print("Brink Admin Panel is already up to date.")
    except:
        print("Error Retrieving Latest Brink Admin Panel")


def auto_select_vig_agent(VIG_TERM, CFG_PATH, site_no):
    term = VIG_TERM
    if len(term) == 2:
        term = term[0] + "EG" + term[1]
    else:
        term = term.upper()
        #print(term)
        #test = input("Press enter to continue")
        if term == 'GRILL':
            term = 'GRILL21'
        if term == 'GRILL 2':
            term = 'GRILL222'
        if term == 'MAKE':
            term = 'MAKE23'
        if term == 'CUSTARD':
            term = 'CUSTARD24'
        if term == 'DT EXPO':
            term = 'DTEXPO25'
        if term == 'DT GRILL':
            term = 'DTGRILL26'
        if term == 'DT GRILL 2':
            term = 'DTGRILL227'
        if term == 'DT MAKE':
            term = 'DTMAKE28'
        if term == 'EXPO':
            term = 'EXPO29'
        if term == 'DT EXPO 2':
            term = 'DTEXPO230'

    site_dir = CFG_PATH / site_no
    #print(site_dir)
    if not site_dir.is_dir():
        return None
    agents = []
    fagents = []
    files = os.listdir(site_dir)
    for file in files:
        if file[-4:] == ".exe":
            agents.append(file)
    for x in agents:
        regex = re.search(rf"({re.escape(term)})", x.upper())
        if regex is not None:
            fagents.append(x)
    if len(fagents) == 1:
        print("Auto selecting Vigilix Agent")
        for x in fagents:
            print("Match found execute {}".format(x))
            vigexe = str(site_dir) + "\\" + x
            shutil.copy(vigexe, str(DESKTOP / os.path.basename(x)))
            subprocess.Popen(str(DESKTOP / os.path.basename(x)))
            sleep(5)
            pyautogui.press('enter')
            sleep(10)
            pyautogui.press('space')
            pyautogui.press('tab')
            pyautogui.press('enter')
            sleep(3)
            pyautogui.press('enter')
            return True
    elif len(fagents) > 1:
        while True:
            print("Select Vigilix Agent")
            for num, x in enumerate(fagents):
                print("  {}. {}".format(num + 1, x))
            inputvig = input(">>: ")
            selection = int(inputvig)
            if selection > len(fagents) or selection <= 0:
                print("Invalid selection, please try again")
            else:
                break
        agent = fagents[selection - 1]
        x = str(agent)
        vigexe = str(site_dir) + "\\" + x
        shutil.copy(vigexe, str(DESKTOP / os.path.basename(x)))
        subprocess.Popen(str(DESKTOP / os.path.basename(x)))
        sleep(5)
        pyautogui.press('enter')
        sleep(10)
        pyautogui.press('space')
        pyautogui.press('tab')
        pyautogui.press('enter')
        sleep(3)
        pyautogui.press('enter')
        return True
    else:
        return False


    return False

def post_cleanup(NRO):
    # files that will be deleted
    filelist = [
        r'C:\Users\ws\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\FreddysSetup.bat - Shortcut.lnk',
        r'C:\Users\ws\Desktop\Startup - Shortcut.lnk',
        r'C:\Users\ws\Desktop\README -pdill.txt',
        r'C:\Users\ws\Desktop\FreddysSetup.bat',
        r'C:\Users\Brink\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\FreddysSetup - Shortcut.lnk',
        r'C:\Users\Brink\Desktop\FreddysSetup.bat',
        r'C:\Users\POS\Desktop\FreddysSetup.bat',
        r'C:\Users\POS\Desktop\Startup - Shortcut.lnk',
        r'C:\Users\POS\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\FreddysSetup - Shortcut.lnk'
    ]
    if not NRO:
        print("Run file cleanup? (Deletes unneeded shortcuts etc.)")
        x = input(">>: ")
        x = x.upper()
        if x[0] == "Y":

            for file in filelist:
                if os.path.exists(file):
                    os.remove(file)
                    print("Deleted {}".format(file))
            sleep(2)
            return 0
        else:
            return 0
    else:
        for file in filelist:
            if os.path.exists(file):
                os.remove(file)
                print("Deleted {}".format(file))
        print("Rebooting in 10 seconds...")
        sleep(10)
        os.system("shutdown /r /t 1")
        return 0

# COMMAND

@click.command()
def main():
    # retrieve latest Brink Admin Panel
    NRO = False
    get_latest_brinkadminpanel()

    site_no = click.prompt("Site No.?")
    if site_no[0].upper() == "N":
        NRO = True
        print("Running script for a NRO.")
        site_no = site_no[1:]
    kind = menu("Is this a Register or Kitchen?", ["Register", "Kitchen"], True)

    # check if site exists in freddy's configs &
    # get list of registers and kitchens

    site = get_site_info(site_no, sort=kind == 1)
    print(flush=True)
    if site is None:
        click.echo("Site not found on NAS. Check Internet", err=True)
        input()
        return

    # get specific term/kitchen

    cfg_fn = ["Register.cfg", "Kitchen.cfg"][kind]

    term_id = menu(
        f"Which {['register', 'kitchen'][kind]}?",
        [site["registers"], site["kitchens"]][kind],
        True
    )
    term_name = [site["registers"], site["kitchens"]][kind][term_id]

    flags = Flags(site_no, term_name)

    # delete sdf

    if Path("C:/Brink/Pos/Register.sdf").is_file():
        if click.confirm("Delete old SDF?", default=True):
            subprocess.run(["taskkill", "/f", "/im", "Register.exe"],
                           stdout=subprocess.DEVNULL)
            sleep(0.5)
            os.remove("C:/Brink/Pos/Register.sdf")

    # cfg path for selected terminal
    cfg_path = str(CFG_PATH / site_no / term_name / cfg_fn)



    # generate pc name
    
    stripped = "".join(term_name.split(" "))
    VIG_TERM = term_name
    pc_name = f"FFC-{site_no}-{stripped}".upper()
    
    click.echo(f"Setting PC name to {pc_name} ...")
    rename_pc(pc_name)

    flags["renamed"] = True

    # get term number
    
    with open(cfg_path) as file:
        m = re.search(r'TerminalNumber="(\d+)"', file.read())
        if m is None:
            click.echo("Couldn't find terminal number", err=True)
            input()
            return

        term_no = int(m.group(1))

    # generate IP only if term 1
    if term_name == "R1" and NRO == True:
        new_ip = IP_PREFIX + str(100 + term_no).zfill(3)

        if click.confirm("Set IP?", default=True):
            click.echo(f"Setting IP to {new_ip} ...")
            set_static_ip(new_ip, "255.255.255.0", "192.168.128.1")

            click.echo(f"Waiting 5 seconds ...")
            sleep(5)
    else:
        new_ip = IP_PREFIX + str(100 + term_no).zfill(3)

        if click.confirm("Set IP?", default=True):
            click.echo(f"Setting IP to {new_ip} ...")
            set_static_ip(new_ip, "255.255.255.0", "192.168.128.1")

            click.echo(f"Waiting 5 seconds ...")
            sleep(5)
    # set timezone

    tzdst = flags["_timezone"]  # use timezone set from other terminal if possible

    if tzdst is None:
        tz_idx = menu(
            "Select timezone",
            TZ_LIST,
            True
        )
        tz = TZ_LIST[tz_idx]

        dst = menu(
            "Daylight Savings Time",
            ["Off (Arizona only)", "On"],
            True
        )

        dst = "" if dst else "_dstoff"

        tzdst = tz + dst
        flags["_timezone"] = tzdst
    
    
    flags["timezone"] = tzdst

    click.echo(f"Setting timezone {tzdst} ...")

    subprocess.run(["tzutil", "/s", tzdst])
    subprocess.run(["net", "start", "W32Time"])
    subprocess.run(["w32tm", "/config", "/update"], stdout=subprocess.DEVNULL)
    subprocess.run(["w32tm", "/resync"], stdout=subprocess.DEVNULL)

    click.echo("Disabling auto-timezone ...")
    
    subprocess.run(["reg", "add",
                    "HKLM\\SYSTEM\\CurrentControlSet\\Services\\tzautoupdate",
                    "/v", "Start", "/t", "REG_DWORD", "/d", "4", "/f"])
    
    print("")  # newline


    # drop config

    try:
        print(f"Dropping {cfg_fn} .", end="")

        new_cfg_dir = "C:\\Brink\\Pos\\" + cfg_fn

        with open(new_cfg_dir, "w") as file:
            file.write(open(cfg_path).read())
        
        print("..")

    except OSError:
        click.echo("Coulnd't find CFG on NAS. Check Internet?")
        input()
        return

    # change sleep setting
    click.echo("Setting sleep to never ...")
    subprocess.run(["powercfg", "-change", "-standby-timeout-ac", "0"])
    subprocess.run(["powercfg", "-change", "-monitor-timeout-ac", "0"])

    print("", flush=True)

    flags["sleep_set"] = True

    # install eBlvd

    eblvd_name = pc_name

    if kind == 1:  # kind = 1 means kitchen
        eblvd_name = f"FFC-{site_no}-K{term_no-20}".upper()
    else:
        if len(eblvd_name) > 16:
            eblvd_name = eblvd_name[:15] + eblvd_name[-1]

    click.echo("Installing eBlvd ...")
    eblvd(eblvd_name)

    flags["eblvd"] = eblvd_name

    #try auto select vigilix install
    #auto_select_vig_agent(VIG_TERM, CFG_PATH, site_no)

    # run vigilix installer
    if not auto_select_vig_agent(VIG_TERM, CFG_PATH, site_no):
        agents = get_vig(site_no)

        agent_idx = menu(
            "Select Vigilix agent",
            [a[0] for a in agents],
            True
        )
        agent = agents[agent_idx]
        exe_path = str(agent[1])

        new_exe_path = DESKTOP / agent[1]

        # copy vig to desktop
        shutil.copy(exe_path, str(DESKTOP / os.path.basename(exe_path)))

        # run vig
        subprocess.run(str(agent[1]))
    
    # delete startup script shortcut
    #file_path = '%userprofile%\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\FreddysSetup.bat - Shortcut.lnk'
    #os.remove(file_path)
    #run cleanup function, deletes startup script etc
    post_cleanup(NRO)
# main 

if __name__ == "__main__":
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    if is_admin():
        quickedit(False)
        main()
    else:
        # Re-run the program with admin rights
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
