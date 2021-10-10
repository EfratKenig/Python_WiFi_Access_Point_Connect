import subprocess
import os
import ctypes, sys


def parse_to_dict(APs_str):
    for outer_ind in range(len(APs_str)):
        APs_str[outer_ind] = (APs_str[outer_ind]).split(' :')
        for inner_ind in range(len(APs_str[outer_ind])):
            APs_str[outer_ind][inner_ind] = APs_str[outer_ind][inner_ind].strip()
            if '    ' in APs_str[outer_ind][inner_ind]:
                new_list = APs_str[outer_ind][inner_ind] = APs_str[outer_ind][inner_ind].split('    ')
                if '' in new_list:
                    new_list.remove('')
    # '\n' means a start of new value in the dictionary
    parsed_dict = {}
    for ap_ind, ap in enumerate(APs_str[:-1]):  # -1 because of the redundant last one
        AP_name = ((APs_str[ap_ind])[1])[0].strip('\n').strip()
        parsed_dict[AP_name] = {}
        for attr_ind, attr in enumerate(APs_str[ap_ind]):
            if isinstance(attr, list):
                # needs to take 2nd item as a new attr in the dictionary
                if isinstance((APs_str[ap_ind])[attr_ind + 1], list):
                    (parsed_dict[AP_name])[attr[1].strip()] = ((APs_str[ap_ind])[attr_ind + 1])[0].strip('\n')
    return parsed_dict


def find_all_APs():
    """runs a command that that shows all availabale access points,
    returns the command output as a list"""
    networks = subprocess.check_output(['netsh', 'wlan', 'show', 'networks', 'mode=Bssid'])
    networks = networks.decode('utf-8')
    networks = networks.replace('\r', '').split("\n\n")[1::]
    return networks


def create_connection(name, password):
    """creates a new connection"""
    connection = """<?xml version=\"1.0\"?>
        <WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
            <name>""" + name + """</name>
            <SSIDConfig>
                <SSID>
                    <name>""" + name + """</name>
                </SSID>
            </SSIDConfig>
            <connectionType>ESS</connectionType>
            <connectionMode>auto</connectionMode>
            <MSM>
                <security>
                    <authEncryption>
                        <authentication>WPA2PSK</authentication>
                        <encryption>AES</encryption>
                        <useOneX>false</useOneX>
                    </authEncryption>
                    <sharedKey>
                        <keyType>passPhrase</keyType>
                        <protected>false</protected>
                        <keyMaterial>""" + password + """</keyMaterial>
                    </sharedKey>
                </security>
            </MSM>
        </WLANProfile>"""
    cmd = "netsh wlan add profile filename=\"" + name + ".xml\""
    with open(name + ".xml", 'w') as xml_file:
        xml_file.write(connection)
    os.system(cmd)


def connect(name):
    cmd = 'netsh wlan connect name="' + name + '"'
    os.system(cmd)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def signal_to_rssi(sig):
    return (sig / 2) - 100




def main():
    """
    flow:
    1: find all available networks
    2: find fastest of all
    3: ask for password if needed
    4: (try to) connect
    5: check status and return
    """
    # find all available networks:
    all_APs_list = find_all_APs()
    for AP in all_APs_list:
        print(AP, end='\n\n')
    # find fastest access point:
    parsed_dict = parse_to_dict(all_APs_list)
    best_rssi = (-100,)
    for key in parsed_dict:
        cur_rssi = signal_to_rssi(int(parsed_dict[key]['Signal'].strip()[:-1]))
        if key == "DIRECT-A9-HP DeskJet 5200 series":
            continue
        if cur_rssi >= best_rssi[0]:
            best_rssi = (cur_rssi, key)
            if cur_rssi == 0:
                break
    print("\n\nBest Access Point:\n"+str(best_rssi[1]))
    for key in parsed_dict[best_rssi[1]]:
        print("\t"+key+":  "+str(parsed_dict[best_rssi[1]][key]))
    if parsed_dict[best_rssi[1]]['Authentication'] != 'Open':
        # ask for password if needed:
        pwd = input("Enter your password\n")
        create_connection(best_rssi[1], pwd)
    if connect(best_rssi[1]) == 0:
        return "Connected Successfully"


if __name__ == "__main__":
    # check admin permissions:
    if is_admin():
        main()
    else:
        # grant permission and run
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        main()
