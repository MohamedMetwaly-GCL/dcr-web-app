"""Exercise the real PowerShell probe with synthetic Windows network inventory."""

import os
from pathlib import Path
import runpy
import subprocess
import unittest


ROOT = Path(__file__).resolve().parent
PROBE = runpy.run_path(str(ROOT / "security_test_support.py"))["_WINDOWS_OFFLINE_PROBE"]
LOOPBACK = "@{InterfaceIndex=1; AddressFamily='IPv4'; ConnectionState='Connected'}"
LOCAL_ROUTE = "@{InterfaceIndex=1; DestinationPrefix='127.0.0.0/8'}"
MINIPORTS = ",".join(
    f"@{{Name='WAN Miniport {index}'; ComponentID='{component}'; HardwareInterface=$false; Status='Up'; ifIndex={index}}}"
    for index, component in enumerate(("ms_ndiswanip", "ms_ndiswanipv6", "ms_ndiswanbh"), start=3)
)


class WindowsOfflinePreflightTests(unittest.TestCase):
    def probe(self, adapters=MINIPORTS, interfaces=LOOPBACK, routes=LOCAL_ROUTE):
        system_root = os.environ["SystemRoot"]
        powershell = Path(system_root) / "System32/WindowsPowerShell/v1.0/powershell.exe"
        inventory = (
            f"function Get-NetAdapter {{ param([switch]$IncludeHidden); if (!$IncludeHidden) {{ throw 'Missing switch' }}; {adapters} }}\n"
            f"function Get-NetIPInterface {{ param([switch]$IncludeAllCompartments); if (!$IncludeAllCompartments) {{ throw 'Missing switch' }}; {interfaces} }}\n"
            f"function Get-NetRoute {{ param([switch]$IncludeAllCompartments); if (!$IncludeAllCompartments) {{ throw 'Missing switch' }}; {routes} }}\n"
        )
        return subprocess.run(
            [str(powershell), "-NoProfile", "-NonInteractive", "-Command", inventory + PROBE],
            env={"SystemRoot": system_root, "WINDIR": system_root},
            capture_output=True, encoding="utf-8", timeout=20,
        )

    def test_inactive_wan_miniports_do_not_block(self):
        result = self.probe()
        self.assertEqual((result.returncode, result.stdout.strip()), (0, "Offline preflight passed"))

    def test_up_tether_ethernet_wifi_and_vpn_block_without_a_default_route(self):
        for name, description, hardware in (
            ("USB tether", "Remote NDIS based Internet Sharing Device", "$true"),
            ("Ethernet", "Ethernet controller", "$true"),
            ("Wi-Fi", "Wireless adapter", "$true"),
            ("VPN", "VPN tunnel adapter", "$false"),
        ):
            with self.subTest(adapter=name):
                adapter = f"@{{Name='{name}'; InterfaceDescription='{description}'; HardwareInterface={hardware}; Status='Up'; ifIndex=20}}"
                result = self.probe(adapters=f"{MINIPORTS},{adapter}")
                self.assertEqual(result.returncode, 2, result.stderr)
                self.assertIn(f"Adapter '{name}' ({description}", result.stdout)

    def test_connected_ip_interface_blocks_even_for_a_miniport(self):
        interface = "@{InterfaceIndex=3; InterfaceAlias='VPN IP'; AddressFamily='IPv4'; CompartmentId=8; ConnectionState='Connected'}"
        result = self.probe(interfaces=f"{LOOPBACK},{interface}")
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn("IP interface 'VPN IP' (IPv4, ifIndex 3, compartment 8)", result.stdout)

    def test_default_routes_in_alternate_compartments_and_miniport_routes_block(self):
        for prefix, index in (("0.0.0.0/0", 20), ("::/0", 20), ("192.0.2.0/24", 3)):
            with self.subTest(prefix=prefix):
                route = f"@{{InterfaceIndex={index}; InterfaceAlias='Tunnel'; DestinationPrefix='{prefix}'; CompartmentId=8}}"
                result = self.probe(routes=f"{LOCAL_ROUTE},{route}")
                self.assertEqual(result.returncode, 2, result.stderr)
                self.assertIn(f"Route '{prefix}' via 'Tunnel' (ifIndex {index}, compartment 8", result.stdout)

    def test_missing_inventory_and_query_errors_fail_closed(self):
        for routes in ("", "throw 'Route query refused'"):
            with self.subTest(routes=routes):
                result = self.probe(routes=routes)
                self.assertEqual(result.returncode, 3, result.stderr)
                self.assertIn("Network inventory failed:", result.stdout)

    def test_miniport_name_without_verified_component_id_blocks(self):
        result = self.probe(adapters="@{Name='WAN Miniport (IP)'; Status='Up'; HardwareInterface=$false; ifIndex=3}")
        self.assertEqual(result.returncode, 2, result.stderr)
        self.assertIn("Adapter 'WAN Miniport (IP)'", result.stdout)


if __name__ == "__main__":
    unittest.main()
