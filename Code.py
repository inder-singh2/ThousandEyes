import paramiko
import openpyxl
import time
from concurrent.futures import ThreadPoolExecutor

# Function to execute commands on a single switch
def configure_switch(ip, username, password, token, guest_ip, vlan, netmask, default_gateway, delay=2):
    try:
        print(f"\nConnecting to switch {ip}...")
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ip, username=username, password=password, timeout=30)
        
        # Start a shell session
        channel = ssh.invoke_shell()
        print(f"Connected to {ip}. Executing commands...")

        # Commands to configure the switch
        commands = [
            "configure terminal",
            "iox",
            "end",
            f"app-hosting install appid TE package flash:thousandeyes-enterprise-agent-5.0.1.cisco.tar",
            "configure terminal",
            "interface GigabitEthernet1/0/13",
            "description Uplink MGMT",
            "switchport access vlan 901",
            "interface AppGigabitEthernet1/0/1",
            "switchport trunk allowed vlan 901",
            "switchport mode trunk",
            f"app-hosting appid TE",
            "app-vnic AppGigabitEthernet trunk",
            f"vlan {vlan} guest-interface 0",
            f"guest-ipaddress {guest_ip} netmask {netmask}",
            "exit",
            f"app-default-gateway {default_gateway} guest-interface 0",
            "name-server0 8.8.8.8",
            "name-server1 8.8.4.4",
            "app-resource docker",
            "prepend-pkg-opts",
            f"run-opts 1 \"-e TEAGENT_ACCOUNT_TOKEN={token}\"",
            "exit",
            "start",
            "end",
            f"app-hosting activate appid TE",
            f"app-hosting start appid TE",
            "write memory"
        ]

        # Execute each command with delay
        for command in commands:
            channel.send(command + "\n")
            time.sleep(delay)
            while not channel.recv_ready():
                time.sleep(1)
            output = channel.recv(65535).decode()
            print(f"Switch {ip} - Command: {command}\n  Output: {output.strip()}")

        channel.close()
        ssh.close()
        print(f"Configuration completed successfully on {ip}.")
    except Exception as e:
        print(f"Failed to configure switch {ip}. Error: {e}")

# Function to load switch details from an Excel file
def load_switch_details(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
        switches = []

        for row in sheet.iter_rows(min_row=2, values_only=True):
            ip, username, password, token, guest_ip, vlan, netmask, default_gateway = row
            if ip and username and password:
                switches.append((ip, username, password, token, guest_ip, vlan, netmask, default_gateway))

        return switches
    except Exception as e:
        print(f"Failed to load Excel file. Error: {e}")
        return []

# Main function
def main():
    excel_file = "switch_details.xlsx"  # Excel file containing switch details
    switches = load_switch_details(excel_file)

    if switches:
        with ThreadPoolExecutor(max_workers=5) as executor:  # Process switches concurrently
            for switch_details in switches:
                executor.submit(configure_switch, *switch_details)
    else:
        print("No switch details found. Please check your Excel file.")

if __name__ == "__main__":
    main()